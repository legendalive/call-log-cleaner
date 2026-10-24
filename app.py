import streamlit as st
import pandas as pd
import re
import time
from io import BytesIO

# Try importing docx for .docx support
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

st.set_page_config(
    page_title="Isolo Imbi Data Logger",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Isolo Imbi Data Logger")
st.caption("Mambo vulai. Ikiwabamba mlete za kabej.")

# Standard default ignored keys per specification
DEFAULT_IGNORED = [
    'tl', 'tl name', 'team leader',
    'ext', 'extension',
    '3cx', '3cx extension', '3cx ext',
    'dcst', 'ccik',
    'employee code', 'emp code', 'agent id'
]

# Initialize session state so data and state persist cleanly across downloads and reruns
if "batch_id" not in st.session_state:
    st.session_state["batch_id"] = 0
if "processed_df" not in st.session_state:
    st.session_state["processed_df"] = None
if "stripped_count" not in st.session_state:
    st.session_state["stripped_count"] = 0

def clear_current_batch():
    """Resets everything: uploaded file, text area, and parsed results."""
    st.session_state["batch_id"] += 1
    st.session_state["processed_df"] = None
    st.session_state["stripped_count"] = 0

def strip_chat_noise(raw_text):
    """
    Strips images, base64 strings, media tags, and chat timestamps
    to reduce memory by 90%+ and prevent crashes on large 1,000+ entries.
    """
    if not raw_text:
        return "", 0

    text = raw_text
    # 1. Strip base64 image strings
    text, b64_count = re.subn(r"data:image/[a-zA-Z0-9\+\-\.]+;base64,[A-Za-z0-9+/=]+", "", text)
    # 2. Strip chat image markers (<image omitted>, [photo], [sticker], etc.)
    img_pattern = r"(?:<image\s+omitted>|\[image(?::\s*[^\]]*)?\]|\[photo(?::\s*[^\]]*)?\]|\[sticker(?::\s*[^\]]*)?\]|\[attached:\s*[^\]]*\]|<attached:\s*[^>]*>|!\[.*?\]\(.*?\)|<img[^>]*>)"
    text, img_count = re.subn(img_pattern, "", text, flags=re.IGNORECASE)
    # 3. Strip audio/video markers
    text = re.sub(r"(?:<audio\s+omitted>|<video\s+omitted>|<document\s+omitted>|\[video(?::\s*[^\]]*)?\]|\[audio(?::\s*[^\]]*)?\])", "", text, flags=re.IGNORECASE)

    # 4. Clean line by line and strip chat headers before field keys
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    cleaned_lines = []
    chat_prefix = re.compile(r"^(?:\[?\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}[,\s]+[\d:]+\s*(?:AM|PM|am|pm)?\]?\s*(?:[-–]\s*)?)(?:[^:]+:\s*)?", re.IGNORECASE)
    simple_prefix = re.compile(r"^\[\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?\]\s*(?:[^:]+:\s*)?", re.IGNORECASE)

    for line in lines:
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue
        if re.match(r"^(?:messages and calls are end-to-end encrypted|security code changed|joined using this group)", stripped, re.IGNORECASE):
            continue
        if chat_prefix.match(stripped):
            sub = chat_prefix.sub("", stripped).strip()
            if re.match(r"^(?:name|issue|start|end|time|ccik|ext|tl|dcst|3cx|employee)", sub, re.IGNORECASE):
                stripped = sub
        elif simple_prefix.match(stripped):
            sub = simple_prefix.sub("", stripped).strip()
            if re.match(r"^(?:name|issue|start|end|time|ccik|ext|tl|dcst|3cx|employee)", sub, re.IGNORECASE):
                stripped = sub
        cleaned_lines.append(stripped)

    return "\n".join(cleaned_lines), (b64_count + img_count)

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
        h = int(std_match.group(1))
        m = int(std_match.group(2))
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

def parse_logs(text, time_mode="Preserve Original", progress_callback=None):
    cleaned_text, stripped_img_count = strip_chat_noise(text)
    lines = cleaned_text.split('\n')
    blocks, current_block = [], []

    for line in lines:
        if re.match(r"^[\$#\s]*name\s*[:=\-\s]", line.strip(), re.IGNORECASE) and any(l.strip() for l in current_block):
            blocks.append(current_block)
            current_block = [line]
        else:
            current_block.append(line)
    if any(l.strip() for l in current_block):
        blocks.append(current_block)

    total_blocks = len(blocks)
    records = []

    for idx, block in enumerate(blocks):
        name = ""
        issue_parts = []
        start_time = ""
        end_time = ""

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

            # 4. Ignored keys (TL, EXT, CCIK, DCST, 3CX, Employee code)
            if is_ignored(stripped):
                continue

            # 5. Explicit Issue
            m_issue = re.match(r"^(?:issue|reason|problem|remarks?|description)\s*[:=\-]?\s*(.+)$", stripped, re.IGNORECASE)
            if m_issue:
                issue_parts.append(m_issue.group(1).strip())
                continue

            # 6. Everything else goes into issues
            issue_parts.append(stripped)

        if name or issue_parts or start_time or end_time:
            records.append({
                "NAME": name,
                "ISSUE": " - ".join(issue_parts) if issue_parts else "",
                "Start time": format_time(start_time, time_mode),
                "END TIME": format_time(end_time, time_mode)
            })

        if progress_callback and (idx % 25 == 0 or idx == total_blocks - 1):
            progress_callback(idx + 1, total_blocks)

    return pd.DataFrame(records, columns=["NAME", "ISSUE", "Start time", "END TIME"]), stripped_img_count

# Layout: Options & Inputs
col_left, col_right = st.columns([2, 1])

with col_right:
    st.subheader("⚙️ Options")
    time_format = st.selectbox(
        "Time Standardization",
        ["Preserve Original", "12-Hour (HH:MM:SS AM/PM)", "24-Hour (HH:MM:SS)"],
        index=0,
        key=f"time_format_{st.session_state['batch_id']}"
    )
    
    st.button(
        "🗑️ Clear Batch",
        on_click=clear_current_batch,
        use_container_width=True,
        help="Clears both uploaded files, pasted text, and table to start a clean new batch"
    )

with col_left:
    st.subheader("📥 Input Data (File or Paste)")
    
    uploaded = st.file_uploader(
        "Upload a .docx, .doc, or .txt file",
        type=["docx", "doc", "txt", "csv", "log"],
        key=f"uploader_{st.session_state['batch_id']}"
    )
    
    file_extracted_text = ""
    if uploaded is not None:
        try:
            if uploaded.name.lower().endswith(".docx"):
                if HAS_DOCX:
                    doc = Document(BytesIO(uploaded.read()))
                    file_extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                    st.success(f"✅ Successfully read `{uploaded.name}`. Images stripped. Click 'Process Batch' below.")
                else:
                    st.warning("⚠️ python-docx not installed. Please paste text directly.")
            else:
                file_extracted_text = uploaded.read().decode("utf-8", errors="ignore")
                st.success(f"✅ Successfully read `{uploaded.name}`. Click 'Process Batch' below.")
        except Exception as e:
            st.error(f"Error reading uploaded file: {e}")

    pasted_text = st.text_area(
        "Or Paste Raw Log / Chat Text Here:",
        value=file_extracted_text,
        key=f"textarea_{st.session_state['batch_id']}",
        height=180,
        placeholder="Paste your unformatted chat dump or call log entries here..."
    )

# The content to process: either from the uploaded file or the text area
content_to_process = pasted_text if pasted_text and pasted_text.strip() else file_extracted_text

# Prominent PROCESS Button requested by user
st.markdown("---")
process_btn_col1, process_btn_col2 = st.columns([1, 2])
with process_btn_col1:
    process_clicked = st.button(
        "⚡ Process Batch",
        type="primary",
        disabled=not bool(content_to_process and content_to_process.strip()),
        use_container_width=True,
        help="Click to parse and clean the logs into the standardized CSV format"
    )

# When Process Button is clicked, run progress bar and save to session_state
if process_clicked and content_to_process.strip():
    progress_bar = st.progress(0, text="Starting parsing engine...")

    def update_progress(current, total):
        pct = min(1.0, current / max(1, total))
        progress_bar.progress(pct, text=f"Processing {current:,} of {total:,} records ({int(pct*100)}%)...")

    # Run parsing with progress updates
    df, stripped_count = parse_logs(content_to_process, time_format, progress_callback=update_progress)
    progress_bar.progress(1.0, text="Processing complete!")
    time.sleep(0.3)
    progress_bar.empty()

    # Save to session_state so results and download button STAY VISIBLE even after clicking download!
    st.session_state["processed_df"] = df
    st.session_state["stripped_count"] = stripped_count

# Results Section: Display processed data and prominent Download Report button
if st.session_state.get("processed_df") is not None and not st.session_state["processed_df"].empty:
    df = st.session_state["processed_df"]
    stripped_count = st.session_state.get("stripped_count", 0)

    st.markdown("### 📊 Cleaned & Standardized Results")
    
    # Prominent, compact Download Button moved UP above the data table
    header_col, action_col = st.columns([3, 1])
    csv_data = df.to_csv(index=False).encode('utf-8')

    with header_col:
        st.write(f"**Total Records:** `{len(df):,}` rows (No caps)")
        if stripped_count > 0:
            st.caption(f"⚡ Automatically stripped {stripped_count} image/media tags to protect memory.")

    with action_col:
        st.download_button(
            label="📥 Download Report",
            data=csv_data,
            file_name="isolo_imbi_call_logs.csv",
            mime="text/csv",
            type="primary",
            use_container_width=False,
            help="Download the standardized 4-column CSV report"
        )

    # Editable Data Preview Table
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
elif st.session_state.get("processed_df") is not None and st.session_state["processed_df"].empty:
    st.warning("No records could be extracted. Please ensure the entries contain lines like 'Name: ...'.")
