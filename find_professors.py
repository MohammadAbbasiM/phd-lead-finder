import os
import re
import time
import html
import requests
import pandas as pd

from datetime import datetime, timezone, timedelta
from difflib import SequenceMatcher

# Optional: used for stronger text-similarity scoring.
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    HAS_SKLEARN = True
except ImportError:
    HAS_SKLEARN = False


# ==============================================================================
# CONFIG
# ==============================================================================

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
OPENALEX_API_KEY = os.environ.get("OPENALEX_API_KEY")  # Optional.

# Detect whichever tracking file is present in the repository.
if os.path.exists("List of emails.xlsx"):
    FILE_NAME = "List of emails.xlsx"
elif os.path.exists("List of emails.xlsx - Sheet1.csv"):
    FILE_NAME = "List of emails.xlsx - Sheet1.csv"
else:
    FILE_NAME = "List of emails.xlsx"

BATCH_SIZE = 5
MAX_FACULTY_PER_UNIVERSITY = 5

# Recent publication window.
RECENT_YEARS = 5

# Search breadth.
WORKS_PER_GROUP = 15
MAX_CANDIDATE_AUTHORS = 20
MAX_WORKS_FOR_FIT = 10

# Optional homepage crawl for funding evidence.
ENABLE_FUNDING_PAGE_CHECK = True
FUNDING_PAGE_TIMEOUT = 12
MAX_HOMEPAGE_CHARS = 50000

OPENALEX_BASE = "https://api.openalex.org"
OPENALEX_EMAIL = "phd-lead-finder@academic.org"

# ==============================================================================
# CANDIDATE RESEARCH PROFILE
# ==============================================================================

# Strong terms receive much larger weights than generic AI/ML/robotics terms.
# This is deliberately asymmetric: a professor doing VIO should outrank
# someone doing generic machine learning.
SEARCH_GROUPS = [
    (
        "Visual-Inertial Navigation & SLAM",
        [
            "visual inertial odometry",
            "visual inertial navigation",
            "visual SLAM",
            "visual-inertial SLAM",
            "visual odometry",
            "camera IMU fusion",
            "inertial navigation",
        ],
    ),
    (
        "Robotics & Autonomous Navigation",
        [
            "robot localization",
            "autonomous navigation",
            "mobile robot navigation",
            "robot perception",
            "autonomous robotics",
            "field robotics",
            "outdoor robotics",
        ],
    ),
    (
        "GNSS & Multi-Sensor Navigation",
        [
            "GNSS positioning",
            "GNSS navigation",
            "GNSS INS integration",
            "GNSS sensor fusion",
            "multi-sensor fusion",
            "robust localization",
            "precise positioning",
        ],
    ),
    (
        "Computer Vision & 3D Perception",
        [
            "computer vision robotics",
            "3D computer vision",
            "visual perception",
            "3D perception",
            "visual localization",
            "depth estimation",
            "stereo vision",
        ],
    ),
    (
        "LiDAR & Multimodal Sensor Fusion",
        [
            "LiDAR camera fusion",
            "LiDAR inertial fusion",
            "LiDAR visual odometry",
            "LiDAR SLAM",
            "multimodal sensor fusion",
            "camera LiDAR IMU",
        ],
    ),
    (
        "State Estimation & Optimization",
        [
            "state estimation robotics",
            "pose estimation",
            "trajectory estimation",
            "factor graph optimization",
            "nonlinear optimization robotics",
            "probabilistic robotics",
            "Kalman filtering robotics",
        ],
    ),
    (
        "Machine Learning for Robotics & Perception",
        [
            "deep learning robotics",
            "machine learning robotics",
            "deep learning computer vision",
            "self-supervised visual navigation",
            "learning-based localization",
            "neural visual odometry",
            "learning-based perception",
        ],
    ),
    (
        "Autonomous Vehicles & Intelligent Systems",
        [
            "autonomous vehicles",
            "autonomous driving perception",
            "vehicle localization",
            "intelligent transportation systems",
            "robotic navigation",
            "autonomous systems perception",
        ],
    ),
    (
        "Embedded AI & Edge Robotics",
        [
            "embedded AI robotics",
            "edge AI robotics",
            "real-time computer vision",
            "GPU accelerated robotics",
            "embedded computer vision",
            "AI acceleration robotics",
            "Jetson robotics",
        ],
    ),
    (
        "Medical Image Computing",
        [
            "medical image segmentation",
            "medical image processing",
            "deep learning medical imaging",
            "computer vision medical imaging",
            "image segmentation",
        ],
    ),
    (
        "Embedded Systems & Digital Hardware",
        [
            "embedded systems",
            "embedded AI",
            "FPGA neural network accelerator",
            "hardware acceleration",
            "digital hardware design",
            "real-time embedded systems",
        ],
    ),
]

# Explicitly weighted research concepts.
# Lower-weight generic terms prevent "AI", "robotics", etc. from dominating.
TIER1_PHRASES = {
    "visual inertial odometry": 10.0,
    "visual inertial navigation": 10.0,
    "visual-inertial slam": 10.0,
    "visual slam": 9.5,
    "visual odometry": 9.0,
    "camera imu fusion": 9.0,
    "gnss ins integration": 9.0,
    "gnss sensor fusion": 9.0,
    "gnss positioning": 8.5,
    "robot localization": 8.5,
    "visual localization": 8.5,
    "autonomous navigation": 8.5,
    "factor graph optimization": 8.0,
    "state estimation": 8.0,
    "lidar slam": 8.0,
    "lidar camera fusion": 8.0,
    "camera lidar imu": 8.0,
    "inertial navigation": 8.0,
    "multi-sensor fusion": 8.0,
    "sensor fusion": 7.5,
    "robust localization": 7.5,
    "pose estimation": 7.0,
    "trajectory estimation": 7.0,
    "computer vision robotics": 7.0,
    "robot perception": 7.0,
}

TIER2_PHRASES = {
    "3d computer vision": 5.5,
    "3d perception": 5.5,
    "visual perception": 5.5,
    "stereo vision": 5.0,
    "depth estimation": 5.0,
    "field robotics": 5.0,
    "mobile robot navigation": 5.0,
    "autonomous vehicles": 5.0,
    "autonomous driving perception": 5.0,
    "probabilistic robotics": 5.0,
    "neural visual odometry": 5.0,
    "learning-based localization": 4.5,
    "self-supervised visual navigation": 4.5,
    "learning-based perception": 4.5,
    "lidar inertial fusion": 5.5,
    "multimodal sensor fusion": 5.5,
}

TIER3_PHRASES = {
    "deep learning": 2.5,
    "machine learning": 2.0,
    "computer vision": 2.0,
    "object detection": 2.0,
    "object tracking": 2.0,
    "embedded ai": 2.0,
    "edge ai": 2.0,
    "cuda": 1.5,
    "gpu": 1.5,
    "ros": 1.5,
    "ros2": 1.5,
    "embedded systems": 1.5,
    "fpga": 1.0,
}

RESEARCH_PROFILE_TEXT = " ".join(
    phrase for _, phrases in SEARCH_GROUPS for phrase in phrases
)

# Funding phrases only count as evidence when found on a faculty/homepage page.
FUNDING_HIGH = [
    "fully funded phd",
    "funded phd",
    "funded doctoral",
    "phd studentship",
    "doctoral studentship",
    "phd scholarship",
    "funded position",
    "funded positions",
    "phd opportunity",
    "phd opportunities",
]

FUNDING_MEDIUM = [
    "phd positions",
    "doctoral positions",
    "phd students",
    "doctoral students",
    "research assistantship",
    "research assistant",
    "scholarship",
    "studentship",
]

# ==============================================================================
# HTTP HELPERS
# ==============================================================================

session = requests.Session()
session.headers.update(
    {
        "User-Agent": f"Mozilla/5.0 PhDLeadFinder/2.0 (+mailto:{OPENALEX_EMAIL})",
        "Accept-Language": "en-US,en;q=0.9",
    }
)


def openalex_get(path_or_url, params=None, timeout=20, retries=3):
    """GET an OpenAlex endpoint with retries and useful diagnostics."""
    if path_or_url.startswith("http://") or path_or_url.startswith("https://"):
        url = path_or_url
    else:
        url = f"{OPENALEX_BASE}/{path_or_url.lstrip('/')}"

    params = dict(params or {})
    params.setdefault("mailto", OPENALEX_EMAIL)

    if OPENALEX_API_KEY:
        params.setdefault("api_key", OPENALEX_API_KEY)

    last_error = None

    for attempt in range(retries):
        try:
            response = session.get(url, params=params, timeout=timeout)
            if response.status_code == 200:
                return response.json()

            if response.status_code in (429, 500, 502, 503, 504):
                wait = min(10 * (attempt + 1), 30)
                print(f"    [!] OpenAlex HTTP {response.status_code}; retrying in {wait}s...")
                time.sleep(wait)
                continue

            print(f"    [!] OpenAlex HTTP {response.status_code}: {response.text[:250]}")
            return None

        except Exception as exc:
            last_error = exc
            time.sleep(2 * (attempt + 1))

    print(f"    [!] OpenAlex request failed after retries: {last_error}")
    return None


# ==============================================================================
# TEXT / NORMALIZATION
# ==============================================================================

def normalize_name(value):
    value = str(value or "").lower()
    value = re.sub(r"[^\w\s]", " ", value, flags=re.UNICODE)
    return re.sub(r"\s+", " ", value).strip()


def clean_university_name(name):
    """
    Clean Excel annotations without incorrectly splitting legitimate names such as
    'University of Science and Technology of China'.
    """
    name = str(name).strip()
    name = re.sub(r"\s+", " ", name)
    name = re.sub(r"\([^)]*\)", "", name).strip()

    # Remove common spreadsheet notes after separators.
    name = re.sub(r"\s*-\s*3years?\s+phd\s+programs?.*$", "", name, flags=re.I)
    name = re.sub(r"\s*\(\s*3years?\s+phd programs?\s*\)\s*$", "", name, flags=re.I)

    # "ETH Zürich and Bologna" is not a single OpenAlex institution.
    # In that special case, search the first institution.
    if normalize_name(name) == normalize_name("ETH Zürich and Bologna"):
        name = "ETH Zürich"

    return name.strip(" -")


def reconstruct_abstract(inverted_index):
    if not inverted_index:
        return ""

    positions = []
    for word, indexes in inverted_index.items():
        for idx in indexes:
            positions.append((idx, word))

    positions.sort(key=lambda x: x[0])
    return " ".join(word for _, word in positions)


def work_text(work):
    parts = [
        work.get("title", ""),
        reconstruct_abstract(work.get("abstract_inverted_index")),
    ]

    for topic in work.get("topics") or []:
        topic_name = topic.get("display_name", "")
        if topic_name:
            parts.append(topic_name)

    for keyword in work.get("keywords") or []:
        keyword_name = keyword.get("display_name", "")
        if keyword_name:
            parts.append(keyword_name)

    return " ".join(parts).strip()


def phrase_score(text):
    """
    Weighted phrase matching.
    Tier-1 navigation concepts dominate generic AI/ML keywords.
    """
    lower = normalize_name(text)
    total = 0.0
    matched = []

    all_phrases = [
        ("T1", TIER1_PHRASES),
        ("T2", TIER2_PHRASES),
        ("T3", TIER3_PHRASES),
    ]

    for tier, mapping in all_phrases:
        for phrase, weight in mapping.items():
            normalized_phrase = normalize_name(phrase)
            if normalized_phrase in lower:
                total += weight
                matched.append((phrase, weight, tier))

    # Mild bonus for explicitly strong combinations.
    combo_pairs = [
        ("gnss", "sensor fusion"),
        ("camera", "imu"),
        ("lidar", "slam"),
        ("visual", "navigation"),
        ("robot", "localization"),
        ("factor graph", "navigation"),
    ]
    for a, b in combo_pairs:
        if a in lower and b in lower:
            total += 2.0

    matched.sort(key=lambda x: x[1], reverse=True)
    return total, matched


def tfidf_similarity(profile_text, texts):
    """
    TF-IDF similarity is used as a lightweight semantic-text similarity layer.
    It is optional; if scikit-learn is unavailable, the algorithm falls back
    to weighted phrase matching.
    """
    if not texts:
        return []

    if not HAS_SKLEARN:
        return [0.0 for _ in texts]

    try:
        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2),
            min_df=1,
            max_features=6000,
        )
        matrix = vectorizer.fit_transform([profile_text] + texts)
        sims = cosine_similarity(matrix[0:1], matrix[1:]).flatten()
        return [float(max(0.0, s)) for s in sims]
    except Exception:
        return [0.0 for _ in texts]


def recent_year_start():
    dt = datetime.now(timezone.utc) - timedelta(days=365.25 * RECENT_YEARS)
    return dt.date().isoformat()


def recency_factor(publication_date):
    if not publication_date:
        return 0.2

    try:
        dt = datetime.fromisoformat(publication_date)
    except ValueError:
        return 0.2

    now = datetime.now(timezone.utc)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)

    days = max(0, (now - dt).days)

    if days <= 180:
        return 1.00
    if days <= 365:
        return 0.90
    if days <= 730:
        return 0.78
    if days <= 1095:
        return 0.68
    if days <= 1460:
        return 0.58
    return 0.45


# ==============================================================================
# INSTITUTION RESOLUTION
# ==============================================================================

def institution_name_similarity(query_name, candidate_name):
    q = normalize_name(query_name)
    c = normalize_name(candidate_name)

    q_tokens = set(q.split())
    c_tokens = set(c.split())

    token_score = len(q_tokens & c_tokens) / max(len(q_tokens), 1)
    sequence_score = SequenceMatcher(None, q, c).ratio()

    return 0.65 * token_score + 0.35 * sequence_score


def resolve_institution(raw_uni_name):
    clean_name = clean_university_name(raw_uni_name)
    print(f"--> Resolving OpenAlex institution: {clean_name}")

    data = openalex_get(
        "/institutions",
        params={
            "search": clean_name,
            "per_page": 10,
        },
    )

    if not data or not data.get("results"):
        print(f"    [!] Institution not resolved: {clean_name}")
        return None

    ranked = sorted(
        data["results"],
        key=lambda item: institution_name_similarity(
            clean_name, item.get("display_name", "")
        ),
        reverse=True,
    )

    best = ranked[0]
    similarity = institution_name_similarity(
        clean_name, best.get("display_name", "")
    )

    if similarity < 0.45:
        print(
            f"    [!] Weak institution match: {best.get('display_name')} "
            f"(similarity={similarity:.2f})"
        )
        return None

    print(
        f"    [✓] Matched: {best.get('display_name')} "
        f"({best.get('id')}, similarity={similarity:.2f})"
    )
    return best


# ==============================================================================
# WORK SEARCH + FACULTY AGGREGATION
# ==============================================================================

def build_group_query(keywords):
    return " OR ".join(f'"{keyword}"' for keyword in keywords)


def author_belongs_to_institution(authorship, institution_id):
    target_id = institution_id.rstrip("/").split("/")[-1]
    for inst in authorship.get("institutions") or []:
        inst_id = (inst.get("id") or "").rstrip("/").split("/")[-1]
        if inst_id == target_id:
            return True
    return False


def search_faculty_broad(raw_uni_name):
    """
    Returns:
      {
        "ok": bool,
        "institution": dict,
        "faculty": [faculty_record, ...]
      }
    """
    institution = resolve_institution(raw_uni_name)

    if not institution:
        return {"ok": False, "institution": None, "faculty": []}

    inst_id = institution["id"].rstrip("/").split("/")[-1]
    start_date = recent_year_start()

    all_works = {}
    author_stats = {}

    for field_label, keywords in SEARCH_GROUPS:
        query_str = build_group_query(keywords)

        data = openalex_get(
            "/works",
            params={
                "filter": (
                    f"authorships.institutions.id:{inst_id},"
                    f"from_publication_date:{start_date}"
                ),
                "search": query_str,
                "sort": "relevance_score:desc,publication_date:desc",
                "per_page": WORKS_PER_GROUP,
                # Keep payload reasonably small.
                "select": (
                    "id,title,publication_date,publication_year,"
                    "authorships,abstract_inverted_index,topics,keywords,"
                    "cited_by_count"
                ),
            },
        )

        if not data:
            continue

        for work in data.get("results", []):
            work_id = work.get("id")
            if not work_id:
                continue

            # Keep the work once, but remember every matched group.
            if work_id not in all_works:
                all_works[work_id] = {
                    "work": work,
                    "fields": set(),
                }

            all_works[work_id]["fields"].add(field_label)

    if not all_works:
        return {"ok": True, "institution": institution, "faculty": []}

    work_records = list(all_works.values())
    texts = [work_text(item["work"]) for item in work_records]
    sims = tfidf_similarity(RESEARCH_PROFILE_TEXT, texts)

    for idx, item in enumerate(work_records):
        work = item["work"]
        text = texts[idx]

        weighted_score, matched = phrase_score(text)
        # Normalize weighted phrase score into 0..1.
        # 45 is a deliberately high "excellent direct fit" reference point.
        keyword_norm = min(weighted_score / 45.0, 1.0)

        # Optional TF-IDF similarity. It is a supporting signal, never the
        # dominant signal, so generic ML papers do not rise to the top.
        similarity_norm = min(sims[idx] * 3.0, 1.0)

        base_fit = 0.80 * keyword_norm + 0.20 * similarity_norm
        recency = recency_factor(work.get("publication_date"))

        # More recent matching papers count more.
        contribution = base_fit * (0.70 + 0.30 * recency)

        for authorship in work.get("authorships") or []:
            if not author_belongs_to_institution(authorship, inst_id):
                continue

            author = authorship.get("author") or {}
            author_id = author.get("id")
            author_name = author.get("display_name")

            if not author_id or not author_name:
                continue

            if author_id not in author_stats:
                author_stats[author_id] = {
                    "author_id": author_id,
                    "professor": author_name,
                    "matched_works": {},
                    "fields": {},
                }

            author_stats[author_id]["matched_works"][work.get("id")] = {
                "title": work.get("title", "N/A"),
                "publication_date": work.get("publication_date", ""),
                "year": work.get("publication_year", ""),
                "cited_by_count": work.get("cited_by_count", 0),
                "fit": contribution,
                "raw_fit": base_fit,
                "matched_phrases": [x[0] for x in matched[:8]],
            }

            for field in item["fields"]:
                author_stats[author_id]["fields"][field] = (
                    author_stats[author_id]["fields"].get(field, 0) + contribution
                )

    # Preliminary ranking: strong research fit + breadth.
    for author_id, rec in author_stats.items():
        works = list(rec["matched_works"].values())

        if not works:
            rec["pre_score"] = 0.0
            continue

        total_fit = sum(w["fit"] for w in works)
        best_fit = max(w["raw_fit"] for w in works)
        breadth = len(rec["fields"])

        rec["pre_score"] = (
            min(total_fit / 5.0, 5.0)
            + min(best_fit * 4.0, 4.0)
            + min(breadth * 0.25, 1.0)
        )

    top_candidates = sorted(
        author_stats.values(),
        key=lambda x: x["pre_score"],
        reverse=True,
    )[:MAX_CANDIDATE_AUTHORS]

    # Enrich the strongest candidate authors.
    faculty = []
    for rec in top_candidates:
        author_data = get_author_details(rec["author_id"])
        if not author_data:
            author_data = {}

        recent_author_works = get_recent_author_works(
            rec["author_id"],
            recent_year_start(),
            per_page=25,
        )

        enriched = score_faculty_record(
            raw_uni_name=raw_uni_name,
            institution=institution,
            candidate=rec,
            author_data=author_data,
            recent_author_works=recent_author_works,
        )
        faculty.append(enriched)

        time.sleep(0.15)

    faculty.sort(
        key=lambda x: x["faculty_score"],
        reverse=True,
    )

    return {
        "ok": True,
        "institution": institution,
        "faculty": faculty[:MAX_FACULTY_PER_UNIVERSITY],
    }


def get_author_details(author_id):
    return openalex_get(
        f"/authors/{author_id.rstrip('/').split('/')[-1]}",
        timeout=15,
    )


def get_recent_author_works(author_id, start_date, per_page=25):
    data = openalex_get(
        "/works",
        params={
            "filter": (
                f"author.id:{author_id.rstrip('/').split('/')[-1]},"
                f"from_publication_date:{start_date}"
            ),
            "sort": "publication_date:desc",
            "per_page": per_page,
            "select": (
                "id,title,publication_date,publication_year,cited_by_count,"
                "authorships"
            ),
        },
        timeout=18,
    )
    return (data or {}).get("results", [])


# ==============================================================================
# FACULTY SCORING
# ==============================================================================

def activity_score(recent_works):
    count = len(recent_works)

    if count >= 10:
        count_score = 10.0
    elif count >= 7:
        count_score = 9.0
    elif count >= 5:
        count_score = 8.0
    elif count >= 3:
        count_score = 6.5
    elif count >= 2:
        count_score = 5.0
    elif count == 1:
        count_score = 3.0
    else:
        count_score = 0.0

    latest_date = ""
    if recent_works:
        latest_date = max(
            [w.get("publication_date", "") for w in recent_works if w.get("publication_date", "")],
            default="",
        )

    latest_factor = recency_factor(latest_date)
    latest_score = latest_factor * 10.0

    return round(0.65 * count_score + 0.35 * latest_score, 2), latest_date


def detect_funding_from_homepage(homepage_url):
    if not ENABLE_FUNDING_PAGE_CHECK or not homepage_url:
        return {
            "score": 0.0,
            "evidence": "No homepage/funding evidence available",
            "url": homepage_url or "",
        }

    try:
        response = session.get(
            homepage_url,
            timeout=FUNDING_PAGE_TIMEOUT,
            allow_redirects=True,
        )
        if response.status_code != 200:
            return {
                "score": 0.0,
                "evidence": f"Homepage unavailable (HTTP {response.status_code})",
                "url": homepage_url,
            }

        text = response.text[:MAX_HOMEPAGE_CHARS].lower()
        text = re.sub(r"\s+", " ", text)

        high_hits = [p for p in FUNDING_HIGH if p in text]
        medium_hits = [p for p in FUNDING_MEDIUM if p in text]

        if high_hits:
            return {
                "score": 10.0,
                "evidence": "Funding-related language: " + ", ".join(high_hits[:4]),
                "url": response.url,
            }

        if medium_hits:
            return {
                "score": 6.0,
                "evidence": "PhD/research-student language: " + ", ".join(medium_hits[:4]),
                "url": response.url,
            }

        return {
            "score": 0.0,
            "evidence": "No explicit PhD funding evidence found on homepage",
            "url": response.url,
        }

    except Exception as exc:
        return {
            "score": 0.0,
            "evidence": f"Funding page check failed: {type(exc).__name__}",
            "url": homepage_url,
        }


def score_faculty_record(
    raw_uni_name,
    institution,
    candidate,
    author_data,
    recent_author_works,
):
    matched_works = list(candidate["matched_works"].values())
    matched_works.sort(
        key=lambda w: (
            w.get("fit", 0.0),
            w.get("publication_date", ""),
            w.get("cited_by_count", 0),
        ),
        reverse=True,
    )

    top_fit_values = [w["raw_fit"] for w in matched_works[:MAX_WORKS_FOR_FIT]]

    if top_fit_values:
        # Best paper matters, but repeated high-quality matches matter too.
        best = max(top_fit_values)
        mean_top = sum(top_fit_values) / len(top_fit_values)
        research_score = min((0.60 * best + 0.40 * mean_top) * 10.0, 10.0)
    else:
        research_score = 0.0

    activity, latest_publication = activity_score(recent_author_works)

    homepage_url = author_data.get("homepage_url") or ""

    funding = detect_funding_from_homepage(homepage_url)

    faculty_score = (
        0.60 * research_score
        + 0.20 * activity
        + 0.20 * funding["score"]
    )

    field_scores = candidate["fields"]
    top_fields = [
        field
        for field, _score in sorted(
            field_scores.items(),
            key=lambda x: x[1],
            reverse=True,
        )[:3]
    ]

    top_work = matched_works[0] if matched_works else {}
    recent_title = top_work.get("title", "N/A")
    recent_year = top_work.get("year", "")

    return {
        "field": top_fields[0] if top_fields else "General CV/Robotics",
        "fields": top_fields,
        "professor": candidate["professor"],
        "university": raw_uni_name,
        "author_id": candidate["author_id"],
        "profile_url": author_data.get("id", candidate["author_id"]),
        "homepage_url": homepage_url,
        "paper": recent_title,
        "year": recent_year,
        "research_fit_score": round(research_score, 2),
        "activity_score": round(activity, 2),
        "funding_score": round(funding["score"], 2),
        "faculty_score": round(faculty_score, 2),
        "matched_works_count": len(matched_works),
        "recent_works_count": len(recent_author_works),
        "latest_publication": latest_publication,
        "funding_evidence": funding["evidence"],
        "funding_url": funding["url"],
        "recent_works": matched_works[:3],
    }


# ==============================================================================
# TELEGRAM
# ==============================================================================

def send_telegram_html(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("[!] Missing Telegram credentials.")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }

    try:
        response = requests.post(url, json=payload, timeout=20)

        if response.status_code != 200:
            print(f"[!] Telegram HTTP {response.status_code}: {response.text[:500]}")
            return False

        body = response.json()
        if not body.get("ok"):
            print(f"[!] Telegram API error: {body}")
            return False

        return True

    except Exception as exc:
        print(f"[!] Telegram send error: {exc}")
        return False


def send_telegram_document(file_path, caption=""):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"

    try:
        with open(file_path, "rb") as doc:
            response = requests.post(
                url,
                data={
                    "chat_id": TELEGRAM_CHAT_ID,
                    "caption": caption,
                },
                files={"document": doc},
                timeout=35,
            )

        if response.status_code != 200:
            print(f"[!] Telegram document HTTP {response.status_code}: {response.text[:500]}")
            return False

        body = response.json()
        if not body.get("ok"):
            print(f"[!] Telegram document API error: {body}")
            return False

        return True

    except Exception as exc:
        print(f"[!] Document send error: {exc}")
        return False


# ==============================================================================
# SHEET OUTPUT
# ==============================================================================

OUTPUT_COLUMNS = [
    "Checked_General?",
    "Discovered_Field",
    "General_PI",
    "General_Paper",
    "University_Research_Fit",
    "University_Activity",
    "University_Funding_Evidence",
    "University_Score",
    "Faculty_Breadth",
    "Top_Faculty",
    "Top_Faculty_Details",
    "Funding_Evidence_Details",
    "Last_Publication",
    "Note",
]


def ensure_output_columns(df):
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            if col == "Checked_General?":
                df[col] = "NO"
            else:
                df[col] = ""
    return df


def format_faculty_details(faculty):
    if not faculty:
        return ""

    chunks = []
    for idx, item in enumerate(faculty, start=1):
        chunks.append(
            f"{idx}. {item['professor']} | "
            f"Research={item['research_fit_score']:.1f}/10 | "
            f"Activity={item['activity_score']:.1f}/10 | "
            f"Funding={item['funding_score']:.1f}/10 | "
            f"Overall={item['faculty_score']:.1f}/10 | "
            f"Area={', '.join(item['fields'][:3])} | "
            f"Paper={item['paper'][:100]} ({item['year']})"
        )
    return "\n".join(chunks)


def format_funding_details(faculty):
    if not faculty:
        return ""

    chunks = []
    for item in faculty:
        evidence = item["funding_evidence"]
        if evidence:
            chunks.append(f"{item['professor']}: {evidence}")
    return "\n".join(chunks)


def update_dataframe_row(df, idx, result):
    faculty = result["faculty"]

    if not faculty:
        df.at[idx, "Discovered_Field"] = ""
        df.at[idx, "General_PI"] = ""
        df.at[idx, "General_Paper"] = ""
        df.at[idx, "University_Research_Fit"] = 0
        df.at[idx, "University_Activity"] = 0
        df.at[idx, "University_Funding_Evidence"] = 0
        df.at[idx, "University_Score"] = 0
        df.at[idx, "Faculty_Breadth"] = 0
        df.at[idx, "Top_Faculty"] = ""
        df.at[idx, "Top_Faculty_Details"] = ""
        df.at[idx, "Funding_Evidence_Details"] = ""
        df.at[idx, "Last_Publication"] = ""
        return

    top_n = faculty[:MAX_FACULTY_PER_UNIVERSITY]

    research_score = sum(x["research_fit_score"] for x in top_n) / len(top_n)
    activity_score_avg = sum(x["activity_score"] for x in top_n) / len(top_n)
    funding_score_avg = sum(x["funding_score"] for x in top_n) / len(top_n)

    university_score = (
        0.60 * research_score
        + 0.20 * activity_score_avg
        + 0.20 * funding_score_avg
    )

    breadth = sum(1 for x in top_n if x["research_fit_score"] >= 6.5)

    last_pub = max(
        [x["latest_publication"] for x in top_n if x["latest_publication"]],
        default="",
    )

    first = top_n[0]
    df.at[idx, "Discovered_Field"] = first["field"]
    df.at[idx, "General_PI"] = first["professor"]
    df.at[idx, "General_Paper"] = f"{first['paper']} ({first['year']})"
    df.at[idx, "University_Research_Fit"] = round(research_score, 2)
    df.at[idx, "University_Activity"] = round(activity_score_avg, 2)
    df.at[idx, "University_Funding_Evidence"] = round(funding_score_avg, 2)
    df.at[idx, "University_Score"] = round(university_score, 2)
    df.at[idx, "Faculty_Breadth"] = breadth
    df.at[idx, "Top_Faculty"] = "; ".join(x["professor"] for x in top_n)
    df.at[idx, "Top_Faculty_Details"] = format_faculty_details(top_n)
    df.at[idx, "Funding_Evidence_Details"] = format_funding_details(top_n)
    df.at[idx, "Last_Publication"] = last_pub
    df.at[idx, "Note"] = first["profile_url"]


def save_dataframe(df):
    if FILE_NAME.endswith(".csv"):
        df.to_csv(FILE_NAME, index=False)
    else:
        df.to_excel(FILE_NAME, index=False)


# ==============================================================================
# TELEGRAM REPORT
# ==============================================================================

def build_university_report(universities, all_matches):
    if not all_matches:
        listed = "\n".join(
            f"• {html.escape(u)}" for u in universities
        )
        return (
            "🔎 <b>Faculty Finder v2</b>\n\n"
            "Checked universities:\n"
            f"{listed}\n\n"
            "No high-quality faculty matches were found in this cycle."
        )

    lines = [
        "🎯 <b>Faculty Finder v2 — Top Faculty Matches</b>",
        "",
    ]

    for uni_result in all_matches:
        uni_name = uni_result["university"]
        faculty = uni_result["faculty"]

        if not faculty:
            lines.append(f"🏛 <b>{html.escape(uni_name)}</b> — no strong matches")
            lines.append("")
            continue

        top_n = faculty[:MAX_FACULTY_PER_UNIVERSITY]
        research = sum(x["research_fit_score"] for x in top_n) / len(top_n)
        activity = sum(x["activity_score"] for x in top_n) / len(top_n)
        funding = sum(x["funding_score"] for x in top_n) / len(top_n)
        university_score = (
            0.60 * research + 0.20 * activity + 0.20 * funding
        )

        lines.append(
            f"🏛 <b>{html.escape(uni_name)}</b>\n"
            f"⭐ University score: <b>{university_score:.1f}/10</b>\n"
            f"🎯 Research: {research:.1f} | "
            f"📚 Activity: {activity:.1f} | "
            f"💰 Funding evidence: {funding:.1f}"
        )

        for idx, item in enumerate(top_n, start=1):
            profile_url = item["profile_url"]
            professor = html.escape(item["professor"])
            fields = html.escape(", ".join(item["fields"][:2]))
            paper = html.escape(item["paper"][:90])
            funding_evidence = html.escape(item["funding_evidence"][:130])

            lines.append(
                f"\n{idx}. <b>{professor}</b>\n"
                f"   Fit: {item['research_fit_score']:.1f} | "
                f"Activity: {item['activity_score']:.1f} | "
                f"Funding: {item['funding_score']:.1f} | "
                f"Overall: <b>{item['faculty_score']:.1f}/10</b>\n"
                f"   Area: {fields}\n"
                f"   Paper: <i>{paper}</i>\n"
                f"   Funding evidence: {funding_evidence}\n"
                f"   <a href=\"{html.escape(profile_url)}\">OpenAlex profile</a>"
            )

        lines.append("\n" + "─" * 28)

    message = "\n".join(lines)

    # Telegram message limit is ~4096 characters.
    return message[:3950]


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("🚀 Running Faculty Finder v2...")
    print(f"Loading file: {FILE_NAME}")

    if not os.path.exists(FILE_NAME):
        print(f"[!] Error: {FILE_NAME} not found.")
        return

    try:
        if FILE_NAME.endswith(".csv"):
            df = pd.read_csv(FILE_NAME)
        else:
            df = pd.read_excel(FILE_NAME)
    except Exception as exc:
        print(f"[!] Failed to read tracking file: {exc}")
        return

    if "University" not in df.columns:
        print("[!] Required column 'University' not found.")
        return

    df = ensure_output_columns(df)

    mask = (
        df["University"].notna()
        & (df["University"].astype(str).str.strip() != "")
        & (df["University"].astype(str).str.lower() != "nan")
        & (
            df["Checked_General?"]
            .astype(str)
            .str.upper()
            != "YES"
        )
    )

    pending_rows = df[mask]

    if pending_rows.empty:
        print("[✓] All universities checked.")
        send_telegram_html(
            "🎉 <b>Faculty Finder v2:</b> all universities have been processed."
        )
        return

    batch = pending_rows.head(BATCH_SIZE)
    universities = [str(u).strip() for u in batch["University"].tolist()]

    print(f"Processing batch of {len(universities)}: {universities}")

    all_matches = []

    for idx, row in batch.iterrows():
        raw_uni = str(row["University"]).strip()

        print("\n" + "=" * 72)
        print(f"🏛 {raw_uni}")

        result = search_faculty_broad(raw_uni)

        if not result["ok"]:
            print(f"[!] Search failed for {raw_uni}; leaving row unchecked.")
            continue

        all_matches.append(
            {
                "university": raw_uni,
                "faculty": result["faculty"],
            }
        )

        update_dataframe_row(df, idx, result)
        df.at[idx, "Checked_General?"] = "YES"

        if result["faculty"]:
            best = result["faculty"][0]
            print(
                f"[✓] Best faculty: {best['professor']} | "
                f"Research {best['research_fit_score']:.1f} | "
                f"Activity {best['activity_score']:.1f} | "
                f"Funding {best['funding_score']:.1f} | "
                f"Overall {best['faculty_score']:.1f}"
            )
        else:
            print("[–] No strong faculty match found.")

        time.sleep(0.5)

    # Save progress even if one university failed.
    try:
        save_dataframe(df)
        print(f"[✓] Tracking file updated: {FILE_NAME}")
    except Exception as exc:
        print(f"[!] Failed to save tracking file: {exc}")
        return

    # Telegram report.
    report = build_university_report(universities, all_matches)
    send_telegram_html(report)

    # Send updated sheet.
    send_telegram_document(
        FILE_NAME,
        caption="💾 Updated Faculty Finder v2 tracking sheet.",
    )

    print("✅ Run complete.")


if __name__ == "__main__":
    main()
