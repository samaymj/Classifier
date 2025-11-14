import pandas as pd
import re
import json
from collections import Counter, defaultdict
import string
import os

TICKETS_FILE = "tickets.xlsx"
KEYWORDS_FILE = "category_keywords.xlsx"
OUTPUT_XLSX = "tickets_classified.xlsx"
OUTPUT_JSON = "summary.json"

PREFER_FIRST_MATCH = True
def load_files(tickets_file=TICKETS_FILE, keywords_file=KEYWORDS_FILE):
    tickets = pd.read_excel(tickets_file, engine="openpyxl")
    keywords_df = pd.read_excel(keywords_file, engine="openpyxl")
    return tickets, keywords_df

def clean_text(text):
    if pd.isna(text):
        return ""
    
    text = str(text).lower()
    text = re.sub(rf"[{re.escape(string.punctuation)}]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def build_keyword_map(keywords_df):
    kw_map = {}
    for _, row in keywords_df.iterrows():
        cat = str(row['category']).strip()
        kw_cell = row.get('keywords', "")
        if pd.isna(kw_cell):
            kw_map[cat] = []
            continue
        kws = [k.strip().lower() for k in str(kw_cell).split(",") if k.strip()]
        kw_map[cat] = kws
    return kw_map

def match_ticket(text, kw_map):
    matched = []
    for cat, kws in kw_map.items():
        for kw in kws:
            pattern = r'\b' + re.escape(kw) + r'\b'
            if re.search(pattern, text):
                matched.append(cat)
                break
    return matched

def classify_tickets(tickets_df, kw_map):
    results = []
    counts = Counter()
    unclassified_list = []
    for idx, row in tickets_df.iterrows():
        ticket_id = row.get('ticket_id', f"row_{idx}")
        description = row.get('description', "")
        cleaned = clean_text(description)
        matches = match_ticket(cleaned, kw_map)

        if matches:
            if PREFER_FIRST_MATCH:
                category = matches[0]
            else:
                category = ";".join(matches)
        else:
            category = "Others"
            unclassified_list.append({"ticket_id": ticket_id, "description": description})

        counts[category] += 1
        results.append({
            "ticket_id": ticket_id,
            "description": description,
            "cleaned_description": cleaned,
            "assigned_category": category,
            "matched_categories": matches
        })

    return pd.DataFrame(results), counts, unclassified_list

def save_outputs(classified_df, counts, unclassified_list):
    # Save Excel
    classified_df.to_excel(OUTPUT_XLSX, index=False, engine="openpyxl")

    summary = {
        "total_tickets": int(len(classified_df)),
        "counts_per_category": dict(counts),
        "unclassified_count": int(counts.get("Others", 0)),
        "unclassified_tickets": unclassified_list  # list of dicts
    }

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)

    print(f"Saved classified tickets -> {OUTPUT_XLSX}")
    print(f"Saved summary -> {OUTPUT_JSON}")

def main():
    if not os.path.exists(TICKETS_FILE) or not os.path.exists(KEYWORDS_FILE):
        print(f"Missing files. Make sure {TICKETS_FILE} and {KEYWORDS_FILE} exist in the current folder.")
        return

    tickets_df, keywords_df = load_files()
    kw_map = build_keyword_map(keywords_df)

    classified_df, counts, unclassified_list = classify_tickets(tickets_df, kw_map)
    save_outputs(classified_df, counts, unclassified_list)

    print("\n--- Summary (terminal) ---")
    for cat, cnt in counts.most_common():
        print(f"{cat}: {cnt}")
    print("-------------------------\n")

if __name__ == "__main__":
    main()

