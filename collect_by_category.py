"""
==============================================================
PROTOTYPE 3 — CATEGORY-AWARE MARATHI COMMENT COLLECTOR
==============================================================

Strategy
--------
• Reads words directly from "Lexical rev.xlsx" (Filtered_Lexicon sheet).
• Only uses terms labelled HATE, OFFN, or PRFN; NOT words are excluded.
• Uses content-driven discovery (general Marathi search queries) —
  offensive words are NEVER put in the YouTube search bar.
• Every collected comment is labelled with the highest-priority
  category of the lexicon term that matched it.
  Priority: HATE > OFFN > PRFN
• One output CSV with a 'category' column so the dataset can be
  split or used as-is for annotation.

Output (in output/)
-------------------
  comments_by_category.csv   — main annotated dataset
  videos_collected.csv       — video-level metadata
  collection_log.csv         — per-query stats

Run
---
  python collect_by_category.py

Requirements
------------
  pip install google-api-python-client pandas openpyxl
"""

import hashlib
import os
import re
import sys
import time
import unicodedata
from collections import defaultdict
from datetime import datetime, timezone

import pandas as pd
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Reconfigure stdout to UTF-8 so Devanagari titles print without crashing
# on Windows terminals that default to cp1252.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Load variables from the .env file in the same directory as this script.
# override=True ensures .env always wins over system environment variables.
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_SCRIPT_DIR, ".env"), override=True)

# ============================================================
# CONFIGURATION
# ============================================================

# API key is read from .env (YOUTUBE_API_KEY=...) or the system environment.
# Do NOT hardcode the key in this script — put it in the .env file instead.
API_KEY = os.getenv("YOUTUBE_API_KEY", "")

if not API_KEY:
    raise ValueError(
        "YouTube API key not configured.\n"
        "Add this line to prototype_3/.env:\n"
        "  YOUTUBE_API_KEY=YOUR_KEY_HERE"
    )

# Path to the Excel lexicon (one level above this script)
LEXICON_EXCEL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..",
    "Lexical_rev_Mixed_Marathi_English.xlsx",
)
LEXICON_SHEET = "Filtered_Lexicon"

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# COLLECTION LIMITS
# ============================================================

# Max videos to fetch per discovery query (YouTube API hard max = 50)
MAX_VIDEOS_PER_QUERY = 50

# Max comments to read per video (paginated, no API hard limit)
MAX_COMMENTS_PER_VIDEO = 500

# Max matched comments to keep per video (no cap — take everything)
MAX_CANDIDATES_PER_VIDEO = 500

# Allow channels to contribute many videos for diversity
MAX_VIDEOS_PER_CHANNEL = 50

# Target ceiling — set well above 25 000 to ensure we reach it
MAX_TOTAL_CANDIDATES = 30000

# Whether to collect reply comments (costs extra API quota)
COLLECT_REPLIES = True

# YouTube API settings
REGION_CODE = "IN"
RELEVANCE_LANGUAGE = "mr"

# Polite delay between API requests (seconds)
SLEEP_BETWEEN_REQUESTS = 0.1

# Category priority: lowest number = highest priority
# When a comment matches multiple categories, the highest-priority
# one is stored in the 'category' column.
CATEGORY_PRIORITY = {"HATE": 0, "OFFN": 1, "PRFN": 2}

# ============================================================
# DISCOVERY QUERIES
# ============================================================
# Neutral Marathi topic queries used to find videos with
# active comment sections. Offensive words must NOT be placed
# here — lexicon filtering is applied to the comments instead.

DISCOVERY_QUERIES = [
    # --- News (broad) ---
    "Marathi news",
    "Marathi breaking news",
    "Maharashtra news today",
    "आजच्या मराठी बातम्या",
    "मराठी बातम्या",
    "मराठी ताज्या बातम्या",
    "Maharashtra latest news",
    "Marathi news 2024",
    "Marathi news 2025",
    "Marathi news live",

    # --- City / Local news ---
    "Pune news Marathi",
    "Mumbai news Marathi",
    "Nashik news Marathi",
    "Nagpur news Marathi",
    "Aurangabad news Marathi",
    "Kolhapur news Marathi",
    "Solapur news Marathi",
    "Thane news Marathi",
    "Satara news Marathi",
    "Sangli news Marathi",

    # --- Politics ---
    "Marathi politics",
    "Maharashtra politics",
    "मराठी राजकारण",
    "Marathi political debate",
    "Maharashtra election news",
    "Marathi election discussion",
    "Maharashtra vidhansabha",
    "Marathi BJP news",
    "Marathi Congress news",
    "Shiv Sena Marathi",
    "NCP Marathi news",
    "राज ठाकरे",
    "उद्धव ठाकरे",
    "एकनाथ शिंदे",
    "देवेंद्र फडणवीस",
    "अजित पवार",
    "Marathi political rally",

    # --- Social issues ---
    "Marathi social issues",
    "Maharashtra public discussion",
    "Marathi public opinion",
    "Maharashtra farmer protest",
    "Marathi women issues",
    "Maharashtra reservation debate",
    "Marathi language debate",
    "मराठी अस्मिता",
    "मराठी भाषा",

    # --- Crime / current events ---
    "Marathi crime news",
    "Maharashtra crime",
    "Marathi current affairs",
    "Maharashtra accident news",
    "Marathi police news",
    "Maharashtra scam news",
    "महाराष्ट्र गुन्हेगारी",

    # --- Entertainment ---
    "Marathi entertainment news",
    "Marathi interview",
    "Marathi celebrity news",
    "Marathi actor interview",
    "Marathi actress interview",
    "Marathi movie review",
    "Marathi film news",
    "मराठी सिनेमा",
    "मराठी मालिका",
    "Marathi serial latest episode",
    "Marathi web series",

    # --- Vlog / Comedy ---
    "Marathi vlog",
    "Marathi comedy",
    "Marathi funny video",
    "Marathi viral video",
    "Marathi roast",
    "मराठी विनोद",
    "Marathi prank video",
    "Marathi reaction video",
    "Marathi shorts viral",
    "Marathi reels funny",

    # --- Sports ---
    "Marathi cricket news",
    "Maharashtra cricket",
    "Marathi sports news",
    "Maharashtra IPL",
    "Marathi kabaddi",

    # --- Religion / Culture ---
    "Marathi religious debate",
    "Maharashtra temple news",
    "Marathi devotional",
    "गणेशोत्सव मराठी",
    "नवरात्री मराठी",

    # --- Education / Youth ---
    "Marathi student protest",
    "Maharashtra education news",
    "Marathi youth vlog",
    "MPSC Marathi",
    "Marathi job news",

    # --- Business / Economy ---
    "Maharashtra economy news",
    "Marathi business news",
    "Maharashtra budget news",
    "Marathi startup news",
]

# ============================================================
# BUILD YOUTUBE CLIENT
# ============================================================

youtube = build("youtube", "v3", developerKey=API_KEY)


def _validate_api_key():
    """Quick sanity-check: make one cheap API call to verify the key works."""
    try:
        youtube.search().list(
            part="snippet", q="test", type="video", maxResults=1
        ).execute()
        print(f"API key validated OK (key: {API_KEY[:12]}...)")
    except HttpError as e:
        if e.resp.status == 429:
            raise SystemExit(
                "\n[QUOTA EXHAUSTED] This API key has no quota left today.\n"
                "Quota resets at midnight Pacific Time (~6:30 AM IST).\n"
                "Add a fresh key to prototype_3/.env as:\n"
                "  YOUTUBE_API_KEY=YOUR_NEW_KEY"
            )
        elif e.resp.status in (400, 403):
            raise SystemExit(
                f"\n[INVALID API KEY] The key '{API_KEY[:12]}...' was rejected by YouTube.\n"
                "Steps to fix:\n"
                "  1. Go to https://console.cloud.google.com\n"
                "  2. Create a new project and enable YouTube Data API v3\n"
                "  3. Create an API key under Credentials\n"
                "  4. Paste it in prototype_3/.env as:\n"
                "       YOUTUBE_API_KEY=YOUR_KEY_HERE"
            )
        else:
            raise SystemExit(f"\n[API ERROR] Unexpected error {e.resp.status}: {e}")


# ============================================================
# LOAD LEXICON FROM EXCEL
# ============================================================


def load_lexicon(excel_path, sheet):
    """
    Read the Filtered_Lexicon sheet from the Excel file.

    Returns a list of dicts:
        {
            'term'       : original term string,
            'normalized' : lower-cased/cleaned form used for matching,
            'category'   : 'HATE' | 'OFFN' | 'PRFN',
        }

    Terms labelled 'NOT' or with NaN in the Review column are excluded.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(
            f"Lexicon Excel not found: {excel_path}\n"
            "Make sure 'Lexical rev.xlsx' is in the project root."
        )

    df = pd.read_excel(excel_path, sheet_name=sheet)

    required = {"term", "Review"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Lexicon sheet '{sheet}' is missing columns: {missing}")

    # Keep only the 3 offensive categories
    df = df[df["Review"].isin(["HATE", "OFFN", "PRFN"])].copy()
    df = df.dropna(subset=["term", "Review"])
    df["term"] = df["term"].astype(str).str.strip()
    df = df[df["term"].str.len() > 0]
    df = df.drop_duplicates(subset=["term"])

    lexicon = []
    for _, row in df.iterrows():
        term = row["term"]
        normalized = _normalize_for_matching(term)
        if not normalized:
            continue
        lexicon.append(
            {
                "term": term,
                "normalized": normalized,
                "category": row["Review"],   # 'HATE', 'OFFN', or 'PRFN'
            }
        )

    return lexicon


# ============================================================
# TEXT NORMALIZATION
# ============================================================


def _normalize_text(text):
    """
    Base normalization used for both matching and deduplication.
    Original comment text is always preserved in the output.
    """
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.lower()
    # Remove zero-width / invisible characters
    text = re.sub(r"[\u200b-\u200f\ufeff]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _normalize_for_matching(text):
    """
    Extended normalization: also collapses repeated characters
    to handle informal spellings (e.g. "gaandu" -> "gandu").
    """
    text = _normalize_text(text)
    # Collapse 3+ repeated characters to 1 (conservative)
    text = re.sub(r"(.)\1{2,}", r"\1", text)
    return text


def _comment_hash(text):
    """SHA-256 of normalized text for duplicate detection."""
    normalized = _normalize_text(text)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


# ============================================================
# LEXICON MATCHING
# ============================================================


def find_matches(comment_text, lexicon):
    """
    Return all lexicon entries whose normalized term appears
    as a substring of the normalized comment.
    """
    normalized_comment = _normalize_for_matching(comment_text)
    return [
        entry
        for entry in lexicon
        if entry["normalized"] in normalized_comment
    ]


def primary_category(matches):
    """
    Given a list of matching lexicon entries, return the
    highest-priority category label.
    Priority order: HATE > OFFN > PRFN
    """
    return min(matches, key=lambda m: CATEGORY_PRIORITY[m["category"]])["category"]


# ============================================================
# YOUTUBE API HELPERS
# ============================================================

# Sentinel raised when the daily Search quota is exhausted.
class QuotaExceededError(Exception):
    pass


def search_videos(query):
    """Search YouTube for videos matching a discovery query.

    Raises QuotaExceededError on HTTP 429 so the caller can
    stop cleanly and save whatever has been collected so far.
    """
    try:
        response = (
            youtube.search()
            .list(
                part="snippet",
                q=query,
                type="video",
                maxResults=MAX_VIDEOS_PER_QUERY,
                regionCode=REGION_CODE,
                relevanceLanguage=RELEVANCE_LANGUAGE,
                order="relevance",
            )
            .execute()
        )
        time.sleep(SLEEP_BETWEEN_REQUESTS)
        return response.get("items", [])
    except HttpError as e:
        if e.resp.status == 429:
            raise QuotaExceededError(
                "Daily YouTube Search quota exhausted. "
                "Quota resets at midnight Pacific Time (~6:30 AM IST). "
                "Saving collected data and stopping."
            ) from e
        print(f"  [ERROR] Video search failed for '{query}': {e}")
        return []


def get_comments(video_id):
    """Fetch top-level comments and replies for a video."""
    comments = []
    next_page_token = None

    while len(comments) < MAX_COMMENTS_PER_VIDEO:
        try:
            response = (
                youtube.commentThreads()
                .list(
                    part="snippet,replies",
                    videoId=video_id,
                    maxResults=min(100, MAX_COMMENTS_PER_VIDEO - len(comments)),
                    pageToken=next_page_token,
                    textFormat="plainText",
                )
                .execute()
            )
            time.sleep(SLEEP_BETWEEN_REQUESTS)
        except HttpError as e:
            # commentsDisabled is expected — skip silently
            if "commentsDisabled" in str(e):
                break
            # Any other error: warn and move on
            print(f"  [WARNING] Could not fetch comments for {video_id}: {e}")
            break

        for item in response.get("items", []):
            top_snippet = item["snippet"]["topLevelComment"]["snippet"]
            comments.append(
                {
                    "comment_id": item["snippet"]["topLevelComment"]["id"],
                    "parent_comment_id": "",
                    "comment_text": top_snippet.get("textDisplay", ""),
                    "published_at": top_snippet.get("publishedAt"),
                    "updated_at": top_snippet.get("updatedAt"),
                    "like_count": top_snippet.get("likeCount", 0),
                    "is_reply": False,
                }
            )

            if COLLECT_REPLIES:
                for reply in item.get("replies", {}).get("comments", []):
                    if len(comments) >= MAX_COMMENTS_PER_VIDEO:
                        break
                    rs = reply["snippet"]
                    comments.append(
                        {
                            "comment_id": reply["id"],
                            "parent_comment_id": rs.get("parentId", ""),
                            "comment_text": rs.get("textDisplay", ""),
                            "published_at": rs.get("publishedAt"),
                            "updated_at": rs.get("updatedAt"),
                            "like_count": rs.get("likeCount", 0),
                            "is_reply": True,
                        }
                    )

        next_page_token = response.get("nextPageToken")
        if not next_page_token:
            break

    return comments


# ============================================================
# MAIN COLLECTION LOOP
# ============================================================


def main():
    print("=" * 70)
    print("PROTOTYPE 3 — CATEGORY-AWARE MARATHI COMMENT COLLECTOR")
    print("=" * 70)

    # Validate API key before doing anything else
    _validate_api_key()

    # --- Load lexicon ---
    print(f"\nLoading lexicon from: {os.path.basename(LEXICON_EXCEL)}")
    lexicon = load_lexicon(LEXICON_EXCEL, LEXICON_SHEET)

    category_counts = defaultdict(int)
    for entry in lexicon:
        category_counts[entry["category"]] += 1

    print(f"Lexicon loaded: {len(lexicon)} terms (NOT words excluded)")
    for cat in ["HATE", "OFFN", "PRFN"]:
        print(f"  {cat}: {category_counts[cat]} terms")
    print()

    # --- Collection state ---
    comment_rows = []

    seen_hashes = set()
    channel_video_count = defaultdict(int)
    total_candidates = 0
    collected_by_category = defaultdict(int)

    # --- Discovery loop ---
    quota_hit = False
    try:
        for query in DISCOVERY_QUERIES:
            if total_candidates >= MAX_TOTAL_CANDIDATES:
                print("\n[STOP] Global candidate ceiling reached.")
                break

            print()
            print("=" * 70)
            print(f"QUERY: {query}")
            print("=" * 70)

            videos = search_videos(query)   # raises QuotaExceededError on 429
            print(f"Videos found: {len(videos)}")

            query_candidates = 0

            for video in videos:
                if total_candidates >= MAX_TOTAL_CANDIDATES:
                    break

                vid_id = video["id"].get("videoId")
                if not vid_id:
                    continue   # skip channels/playlists in search results
                snippet = video["snippet"]
                channel_id = snippet.get("channelId", "")
                channel_title = snippet.get("channelTitle", "")
                video_title = snippet.get("title", "")

                # Channel diversity guard
                if channel_video_count[channel_id] >= MAX_VIDEOS_PER_CHANNEL:
                    print(f"  [SKIP] Channel limit reached: {channel_title}")
                    continue
                channel_video_count[channel_id] += 1

                print(f"\n  Video  : {video_title[:75]}")
                print(f"  Channel: {channel_title}")

                comments = get_comments(vid_id)
                print(f"  Comments fetched: {len(comments)}")

                # Match and categorize comments
                video_candidate_count = 0

                for comment in comments:
                    if total_candidates >= MAX_TOTAL_CANDIDATES:
                        break
                    if video_candidate_count >= MAX_CANDIDATES_PER_VIDEO:
                        break

                    text = comment["comment_text"]
                    if not text.strip():
                        continue

                    h = _comment_hash(text)
                    if h in seen_hashes:
                        continue

                    matches = find_matches(text, lexicon)
                    if not matches:
                        continue   # No offensive term found — skip

                    seen_hashes.add(h)
                    cat = primary_category(matches)

                    comment_rows.append(
                        {
                            "comment_text": text,
                            "category":     cat,   # HATE / OFFN / PRFN
                        }
                    )

                    video_candidate_count += 1
                    query_candidates += 1
                    total_candidates += 1
                    collected_by_category[cat] += 1

                print(f"  Matched this video: {video_candidate_count}  |  Total so far: {total_candidates}")

    except QuotaExceededError as qe:
        quota_hit = True
        print()
        print("=" * 70)
        print(f"[QUOTA EXHAUSTED] {qe}")
        print("=" * 70)

    # --- Save output ---
    print()
    print("=" * 70)
    print("SAVING OUTPUT FILES")
    print("=" * 70)

    comments_df = pd.DataFrame(comment_rows)   # columns: comment_text, category

    comments_path = os.path.join(OUTPUT_DIR, "comments_by_category.csv")

    # utf-8-sig so Excel opens Devanagari correctly
    comments_df.to_csv(comments_path, index=False, encoding="utf-8-sig")

    # --- Final report ---
    print()
    print("=" * 70)
    print("COLLECTION COMPLETE")
    print("=" * 70)
    print(f"Total comments   : {len(comments_df)}")
    print()
    print("Comments by category:")
    for cat in ["HATE", "OFFN", "PRFN"]:
        print(f"  {cat:6s}: {collected_by_category.get(cat, 0)}")
    print()
    print("File saved:")
    print(f"  {comments_path}")
    print()
    print("IMPORTANT:")
    print("  Category labels (HATE / OFFN / PRFN) are based on lexicon matching.")
    print("  These are CANDIDATE comments — human annotation is required")
    print("  before using this data for model training.")


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()
