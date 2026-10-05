import os
import re
import time
import requests
import pandas as pd

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

# Detect whichever file is present in the repository
if os.path.exists("List of emails.xlsx"):
    FILE_NAME = "List of emails.xlsx"
elif os.path.exists("List of emails.xlsx - Sheet1.csv"):
    FILE_NAME = "List of emails.xlsx - Sheet1.csv"
else:
    FILE_NAME = "List of emails.xlsx"

# ==============================================================================
# BROADER CV RESEARCH TAXONOMY (NON-LOCALIZATION)
# ==============================================================================
# --- PREVIOUS LOCALIZATION GROUPS (COMMENTED OUT) ---
# SEARCH_GROUPS = [
#     ("Localization", ["ultra-wideband", "indoor localization", "indoor positioning", "channel impulse response", "CSI sensing", "TDoA"]),
#     ("Wireless Sensing", ["wireless localization", "RF sensing", "sensor fusion", "ISAC", "6G localization", "pedestrian dead reckoning"]),
# ]

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
            "inertial navigation"
        ]
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
            "outdoor robotics"
        ]
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
            "precise positioning"
        ]
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
            "stereo vision"
        ]
    ),

    (
        "LiDAR & Multimodal Sensor Fusion",
        [
            "LiDAR camera fusion",
            "LiDAR inertial fusion",
            "LiDAR visual odometry",
            "LiDAR SLAM",
            "multimodal sensor fusion",
            "camera LiDAR IMU"
        ]
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
            "Kalman filtering robotics"
        ]
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
            "learning-based perception"
        ]
    ),

    (
        "Autonomous Vehicles & Intelligent Systems",
        [
            "autonomous vehicles",
            "autonomous driving perception",
            "vehicle localization",
            "intelligent transportation systems",
            "robotic navigation",
            "autonomous systems perception"
        ]
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
            "Jetson robotics"
        ]
    ),

    (
        "Medical Image Computing",
        [
            "medical image segmentation",
            "medical image processing",
            "deep learning medical imaging",
            "computer vision medical imaging",
            "image segmentation"
        ]
    ),

    (
        "Embedded Systems & Digital Hardware",
        [
            "embedded systems",
            "embedded AI",
            "FPGA neural network accelerator",
            "hardware acceleration",
            "digital hardware design",
            "real-time embedded systems"
        ]
    )
]

def send_telegram_html(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("[!] Missing Telegram credentials.")
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True
    }
    requests.post(url, json=payload, timeout=20)

def send_telegram_document(file_path, caption=""):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    try:
        with open(file_path, "rb") as doc:
            requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption}, files={"document": doc}, timeout=35)
    except Exception as e:
        print(f"[!] Document send error: {e}")

def clean_university_name(name):
    name = str(name).strip()
    if " and " in name:
        name = name.split(" and ")[0]
    if " / " in name:
        name = name.split(" / ")[-1]
    name = re.sub(r'\(.*?\)', '', name)
    name = re.sub(r'[",]', '', name)
    return name.strip()

def search_faculty_broad(raw_uni_name):
    clean_name = clean_university_name(raw_uni_name)
    print(f"--> Searching OpenAlex for: {clean_name}")
    discovered = []
    headers = {"User-Agent": "mailto:phd_lead_bot@academic.org"}

    try:
        inst_url = f"https://api.openalex.org/institutions?search={requests.utils.quote(clean_name)}"
        res = requests.get(inst_url, headers=headers, timeout=15).json()
        if not res.get("results"):
            print(f"    [!] Institution not resolved for: {clean_name}")
            return discovered
        inst_id = res["results"][0]["id"]
    except Exception as e:
        print(f"    [!] Error looking up institution: {e}")
        return discovered

    for field_label, kw_list in SEARCH_GROUPS:
        query_str = " OR ".join([f'"{k}"' for k in kw_list[:5]])
        works_url = (
            f"https://api.openalex.org/works?"
            f"filter=authorships.institutions.id:{inst_id},from_publication_date:2023-01-01"
            f"&search={requests.utils.quote(query_str)}"
            f"&per-page=2"
        )
        try:
            w_res = requests.get(works_url, headers=headers, timeout=15).json()
            works = w_res.get("results", [])
            for work in works:
                title = work.get("title", "N/A")
                pub_year = work.get("publication_year", "")
                authorships = work.get("authorships", [])
                
                chosen = authorships[-1] if len(authorships) > 1 else (authorships[0] if authorships else None)
                if chosen:
                    author_info = chosen.get("author", {})
                    discovered.append({
                        "field": field_label,
                        "professor": author_info.get("display_name", "Lab PI"),
                        "university": raw_uni_name,
                        "paper": title,
                        "year": pub_year,
                        "profile_url": author_info.get("id", "")
                    })
            if len(discovered) >= 2:
                break
        except Exception as e:
            print(f"    [!] Error querying works for {field_label}: {e}")

    return discovered

def main():
    print(f"Loading file: {FILE_NAME}")
    if not os.path.exists(FILE_NAME):
        print(f"[!] Error: {FILE_NAME} not found.")
        return

    df = pd.read_csv(FILE_NAME) if FILE_NAME.endswith(".csv") else pd.read_excel(FILE_NAME)

    # Initialize new columns if they do not exist
    for col in ["Checked_General?", "Discovered_Field", "General_PI", "General_Paper"]:
        if col not in df.columns:
            df[col] = "NO" if col == "Checked_General?" else ""

    # Select universities not yet scanned for general fields
    mask = (
        (df["University"].notna()) & 
        (df["University"].astype(str).str.strip() != "") & 
        (df["University"].astype(str).str.lower() != "nan") &
        (df["Checked_General?"].astype(str).str.upper() != "YES")
    )
    pending_rows = df[mask]

    if pending_rows.empty:
        print("[✓] All universities checked.")
        send_telegram_html("🎉 <b>All universities have been processed across broader CV fields!</b>")
        return

    batch = pending_rows.head(5)
    universities = batch["University"].tolist()
    print(f"Processing Batch of 5: {universities}")

    all_matches = []
    for idx, row in batch.iterrows():
        uni = str(row["University"]).strip()
        profs = search_faculty_broad(uni)
        df.at[idx, "Checked_General?"] = "YES"

        if profs:
            all_matches.extend(profs)
            df.at[idx, "Discovered_Field"] = profs[0]["field"]
            df.at[idx, "General_PI"] = profs[0]["professor"]
            df.at[idx, "General_Paper"] = f"{profs[0]['paper']} ({profs[0]['year']})"
            df.at[idx, "Note"] = profs[0]["profile_url"]
        time.sleep(1)

    # Save progress back to disk
    if FILE_NAME.endswith(".csv"):
        df.to_csv(FILE_NAME, index=False)
    else:
        df.to_excel(FILE_NAME, index=False)

    # Send report to Telegram
    if all_matches:
        msg = f"🎯 <b>Discovered {len(all_matches)} Labs in Broader CV Fields:</b>\n\n"
        for item in all_matches:
            msg += (
                f"🏛 <b>{item['university']}</b>\n"
                f"🏷 <b>Field:</b> <code>{item['field']}</code>\n"
                f"👤 <b>PI:</b> {item['professor']}\n"
                f"📄 <b>Recent Work ({item['year']}):</b> <i>{item['paper'][:85]}...</i>\n"
                f"🔗 <a href='{item['profile_url']}'>OpenAlex Author Profile</a>\n\n"
            )
        send_telegram_html(msg)
    else:
        send_telegram_html(
            f"🔍 <b>Checked 5 Universities (Broader Fields):</b>\n" + 
            "\n".join([f"• {u}" for u in universities]) + 
            "\n\n<i>No publications found in this cycle.</i>"
        )

    # Send updated spreadsheet directly to Telegram chat
    send_telegram_document(FILE_NAME, caption="💾 Updated tracking sheet (Broader CV Fields).")

if __name__ == "__main__":
    main()
