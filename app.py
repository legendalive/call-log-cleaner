import streamlit as st
import pandas as pd
import re
from io import BytesIO

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

st.set_page_config(page_title="Call Log CSV Standardizer", page_icon="📋", layout="wide")

st.title("📋 Call Log CSV Standardizer")
st.caption("Cleans messy, unstandardized call log entries into standardized 4-column CSV format.")

DEFAULT_IGNORED = [
    'tl', 'tl name', 'team leader',
    'ext', 'extension',
    '3cx', '3cx extension', '3cx ext',
    'dcst', 'ccik',
    'employee code', 'emp code', 'agent id'
]

def format_time(raw_time, mode="Preserve Original"):
    if not raw_time or not str(raw_time).strip():
        return ""
    t = str(raw_time).strip()
    if mode == "Preserve Original":
        return t

    # Military format e.g. 0021hrs or 0118hrs
    mil_match = re.match(r"^([0-2][0-9])([0-5][0-9])\s*(?:hrs?|hours?)?$", t, re.IGNORECASE)
    if mil_match:
        h, m = int(mil_match.group(1)), int(mil_match.group(2))
        if mode == "24-Hour":
            return f"{h:02d}:{m:02d}:00"
        period = "PM" if h >= 12 else "AM"
        dh = 12 if h % 12 == 0 else h % 12
        return f"{dh:02d}:{m:02d}:00 {period}"

    # Standard format e.g. 07:40:57 AM, 1:12 AM
    std_match = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$", t, re.IGNORECASE)
    if std_match:
        h, m = int(std_match.group(1)), int(std_match.group(2))
        s = int(std_match.group(3)) if std_match.group(3) else 0
        meridian = std_match.group(4).upper() if std_match.group(4) else None
        if meridian:
            if meridian == "PM" and h < 12:
                h += 12
            elif meridian == "AM" and h == 12:
                h = 0
        if mode == "24-Hour":
            return f"{h:02d}:{m:02d}:{s:02d}"
        else:
            period = "PM" if h >= 12 else "AM"
            dh = 12 if h % 12 == 0 else h % 12
            return f"{dh:02d}:{m:02d}:{s:02d} {period}"

    return t

def is_ignored(line):
    t = line.strip().lower()
    for key in DEFAULT_IGNORED:
        if re.match(rf"^{re.escape(key)}\s*[:=\-\s]", t) or t == key:
            return True
    return False

def parse_logs(text, time_mode="Preserve Original"):
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    blocks, current_block = [], []

    for line in lines:
        if re.match(r"^[\$#\s]*name\s*[:=\-\s]", line.strip(), re.IGNORECASE) and any(l.strip() for l in current_block):
            blocks.append(current_block)
            current_block = [line]
        else:
            current_block.append(line)
    if any(l.strip() for l in current_block):
        blocks.append(current_block)

    records = []
    for block in blocks:
        name, start_time, end_time = "", "", ""
        issue_parts = []

        for line in block:
            stripped = line.strip()
            if not stripped:
                continue

            # 1. Name
            m_name = re.match(r"^[\$#\s]*name\s*[:=\-]?\s*(.+)$", stripped, re.IGNORECASE)
            if m_name and not name:
                name = m_name.group(1).strip()
                continue

            # 2. Start Time
            m_start = re.match(r"^(?:start\s*time|time\s*in|start)\s*(?:[:=\-]\s*)?(.*)$", stripped, re.IGNORECASE)
            if m_start and not start_time:
                start_time = re.sub(r"^[:=\-]\s*", "", m_start.group(1)).strip()
                continue

            # 3. End Time
            m_end = re.match(r"^(?:end\s*time|time\s*out|end)\s*(?:[:=\-]\s*)?(.*)$", stripped, re.IGNORECASE)
            if m_end and not end_time:
                end_time = re.sub(r"^[:=\-]\s*", "", m_end.group(1)).strip()
                continue

            # 4. Ignored keys
            if is_ignored(stripped):
                continue

            # 5. Explicit Issue
            m_issue = re.match(r"^(?:issue|reason|problem|remarks?|description)\s*[:=\-]?\s*(.+)$", stripped, re.IGNORECASE)
            if m_issue:
                issue_parts.append(m_issue.group(1).strip())
                continue

            # 6. Everything else to issues
            issue_parts.append(stripped)

        if name or issue_parts or start_time or end_time:
            records.append({
                "NAME": name,
                "ISSUE": " - ".join(issue_parts) if issue_parts else "",
                "Start time": format_time(start_time, time_mode),
                "END TIME": format_time(end_time, time_mode)
            })

    return pd.DataFrame(records, columns=["NAME", "ISSUE", "Start time", "END TIME"])

col1, col2 = st.columns([2, 1])
with col2:
    st.subheader("⚙️ Options")
    time_format = st.selectbox("Time Format", ["Preserve Original", "12-Hour (HH:MM:SS AM/PM)", "24-Hour (HH:MM:SS)"])
    if st.button("🗑️ Clear Batch", use_container_width=True):
        st.session_state["raw_input"] = ""
        st.rerun()

with col1:
    st.subheader("📥 Input Raw Logs")
    uploaded = st.file_uploader("Upload .txt or .docx file", type=["txt", "docx"])
    file_text = ""
    if uploaded is not None:
        if uploaded.name.endswith(".docx") and HAS_DOCX:
            doc = Document(BytesIO(uploaded.read()))
            file_text = "\n".join([p.text for p in doc.paragraphs])
        else:
            file_text = uploaded.read().decode("utf-8", errors="ignore")

    input_text = st.text_area("Or Paste Call Log Batch:", value=file_text or st.session_state.get("raw_input", ""), height=220)

if input_text.strip():
    df = parse_logs(input_text, time_format)
    st.subheader(f"Cleaned Records ({len(df)} rows)")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    st.download_button(
        label="📥 Download Standardized CSV",
        data=edited_df.to_csv(index=False).encode('utf-8'),
        file_name="standardized_call_logs.csv",
        mime="text/csv",
        type="primary",
        use_container_width=True
    )
