import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="Mention Processor", layout="centered")
st.title("Mention File Processor")

# --- Inputs ---
uploaded_file = st.file_uploader("Upload the raw Excel file", type=["xlsx"])
file_name = st.text_input("Name of the File (used in Mention ID)", value="")

run_button = st.button("Process File")

# --- Regex pattern (same as notebook) ---
pattern = re.compile(
    r'\b(?:mutual\s*fund|stocks?|shares?|nifty|sensex|trade|mutualfund|gdp|index|nse|bse|equity)\b'
    r'|शेयर|સેન્સેક્સ|मार्केट|निफ्टी|सेंसेक्स|स्टॉक्स',
    re.IGNORECASE
)

def find_stock_terms(x):
    if pd.isna(x):
        return ''
    return ', '.join(pattern.findall(str(x)))

# --- Relevance rules ---

# Which bank a file belongs to, detected from the file name the user types in.
# Add more keywords here if you process other banks later.
BANK_KEYWORDS = ['hdfc', 'icici', 'kotak']

# Nickname to exclude, per bank — only excluded when the file actually belongs
# to that bank (e.g. HDFCBank_Cares is only marked N in HDFC files).
NICKNAME_EXCLUDE = {
    'hdfc': 'hdfcbank_cares',
    'icici': 'icicibank_care',
    'kotak': 'kotakcares',
}

# Competitor brand names that look like HDFC Bank but aren't — only checked
# when processing an HDFC file.
HDFC_COMPETITOR_TERMS = ['hdfclife', 'hdfcergo', 'hdfc ltd.']
ICICI_COMPETITOR_TERMS = ['icicilombard', 'icici lombard']

ALLOWED_RESOURCE_TYPES = {'mass media', 'social network'}
ALLOWED_SOCIAL_SOURCES = {'facebook.com', 'x.com'}


def detect_bank(file_name):
    """Return which bank this file belongs to, based on the typed file name."""
    fn = file_name.lower()
    for kw in BANK_KEYWORDS:
        if kw in fn:
            return kw
    return None


def find_column(columns, target_name):
    """Find a column by name, ignoring case and leading/trailing spaces —
    so a header like ' Resource Type' or 'resource type' still matches
    'Resource type'. Returns the actual column name, or None if not found."""
    target = target_name.strip().lower()
    for c in columns:
        if isinstance(c, str) and c.strip().lower() == target:
            return c
    return None


def compute_relevant(row, bank_keyword, resource_col, nickname_col, source_col):
    resource_type = str(row.get(resource_col, '')).strip().lower() if resource_col else ''
    nickname = str(row.get(nickname_col, '')).strip().lower() if nickname_col else ''
    source = str(row.get(source_col, '')).strip().lower() if source_col else ''
    combine_lower = str(row.get('Combine text', '')).lower()
    stock_val = str(row.get('Stock', '')).strip()

    # Rule 1: Stock column is not blank — a stock/market term was already
    # flagged in the text, so this mention is out of scope.
    if stock_val != '':
        return 'N'

    # Rule 2: only Mass media / Social network are relevant at all — everything
    # else (Forum, Blog, etc.) is marked N immediately.
    if resource_type not in ALLOWED_RESOURCE_TYPES:
        return 'N'

    # Rule 3: for Social network rows, only keep them if Source is Facebook.com
    # or x.com — anything else on social network is marked N.
    if resource_type == 'social network' and source not in ALLOWED_SOCIAL_SOURCES:
        return 'N'

    # Rule 4: bank's own official "care" handle is noise, not a real mention —
    # only applies to the matching bank's file.
    if bank_keyword in NICKNAME_EXCLUDE and nickname == NICKNAME_EXCLUDE[bank_keyword]:
        return 'N'

    # Rule 5: competitor brands that look like HDFC or ICICI but aren't (HDFC Life,
    # HDFC Ergo, HDFC Ltd., ICICI Lombard) — only checked for HDFC files, and
    # only when the text doesn't also mention "bank" (which would mean it's
    # a genuine HDFC Bank mention alongside the other brand name).
    if bank_keyword == 'hdfc':
        if any(term in combine_lower for term in HDFC_COMPETITOR_TERMS) and 'bank' not in combine_lower:
            return 'N'
        
    if bank_keyword == 'icici':
            if any(term in combine_lower for term in ICICI_COMPETITOR_TERMS) and 'bank' not in combine_lower:
                return 'N'

    return ''

FINAL_COLUMNS = [
    'Mention ID', 'Date', 'Time', 'Saved at', 'Title', 'Text', 'Combine text',
    'Sentiment', 'Cluster', 'Stock', 'Relevant', 'Post type',
    'Content Types', 'Source specific format', 'URL', 'Author', 'Nickname',
    'Profile', 'Subscribers', 'Demography', 'Age', 'Source',
    'Publication place', 'Publication place profile',
    'Publication place subscribers', 'Publication place rating',
    'Resource type', 'Language', 'Country', 'Regions', 'City', 'Notes',
    'Reactions', 'Engagement', 'Likes', 'Love', 'Haha', 'Wow', 'Sad',
    'Angry', 'Care', 'Dislikes', 'Comments', 'Reposts', 'Views',
    'Impressions (owned posts)', 'Reach (owned posts)', 'Saves',
    'Potential reach', 'Rating', 'Image brands', 'Image objects',
    'Image scenes', 'Image people', 'Image activities', 'Image type',
    'Image subtype', 'Image URL', 'Assigned to', 'Processed', 'Aspects',
    'Subjects', 'Auto-categories', 'Trend', 'Tags',
]

if run_button:
    if uploaded_file is None:
        st.error("Please upload an Excel file first.")
        st.stop()
    if not file_name.strip():
        st.error("Please enter the Name of the File — it's needed for the Mention ID column.")
        st.stop()

    with st.spinner("Processing..."):
        # Load workbook from the uploaded file
        wb = load_workbook(uploaded_file)

        if 'Mentions' not in wb.sheetnames:
            st.error("This file has no 'Mentions' sheet. Check the upload.")
            st.stop()

        ws = wb['Mentions']

        # Convert to DataFrame
        data = list(ws.values)
        columns = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=columns)


        # Add new columns
        df['Combine text'] = df['Title'].astype(str) + ' ' + df['Text'].astype(str)
        df['Sentiment'] = ''
        df['Cluster'] = ''
        df['Stock'] = ''
        df['Relevant'] = ''

        # Run regex tagging
        df['Stock'] = df['Combine text'].apply(find_stock_terms)

        # Mark Relevant = 'N' per the rules above
        bank_keyword = detect_bank(file_name)
        resource_col = find_column(df.columns, 'Resource type')
        nickname_col = find_column(df.columns, 'Nickname')
        source_col = find_column(df.columns, 'Source')
        relevant_missing_cols = [
            name for name, col in [
                ('Resource type', resource_col),
                ('Nickname', nickname_col),
                ('Source', source_col),
            ] if col is None
        ]
        df['Relevant'] = df.apply(
            lambda row: compute_relevant(row, bank_keyword, resource_col, nickname_col, source_col),
            axis=1
        )

        # Reorder columns — only keep the ones that actually exist in this file
        available_cols = [c for c in FINAL_COLUMNS if c in df.columns]
        missing_cols = [c for c in FINAL_COLUMNS if c not in df.columns]
        df = df[available_cols]

        # Build Mention ID
        current_dt = datetime.now().strftime("%Y-%m-%d")
        df['Mention ID'] = [
            f"{current_dt}/Asha/{file_name}/{i}"
            for i in range(1, len(df) + 1)
        ]

        # Re-order so Mention ID is first (matches notebook's final layout)
        cols = ['Mention ID'] + [c for c in df.columns if c != 'Mention ID']
        df = df[cols]

    if missing_cols:
        st.warning(f"These expected columns were not found in the uploaded file and were skipped: {', '.join(missing_cols)}")

    if bank_keyword is None:
        st.info("Couldn't detect a bank (hdfc/icici/kotak) from the file name you entered — the nickname and competitor-brand Relevant rules were skipped for this run.")

    st.success(f"Done — {len(df)} rows processed.")
    st.dataframe(df.head(20))

    # Prepare download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    st.download_button(
        label="Download processed Excel file",
        data=output,
        file_name=f"{file_name}_processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- Master Mention ID / Contents CSV (only rows not marked N) ---
    csv_df = df.loc[df['Relevant'] != 'N', ['Mention ID', 'Combine text']].rename(
        columns={'Mention ID': 'master Mention ID', 'Combine text': 'Contents'}
    )

    csv_bytes = csv_df.to_csv(index=False, sep=',', encoding='utf-8')

    st.download_button(
        label="Download Master Mention ID / Contents CSV",
        data=csv_bytes.encode('utf-8-sig'),
        file_name=f"{file_name}_master_contents.csv",
        mime="text/csv"
    )