import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import fitz  # PyMuPDF
import pandas as pd
import io
import re
import requests
import cloudinary
import cloudinary.uploader
import tempfile
import os
from datetime import datetime
# =========================================================
# GSO FINDER - CLOUDINARY + FIRESTORE VERSION
# Added:
# 1) CCR No. extraction + storage in Firestore
# 2) PDF storage in Cloudinary instead of Firebase Storage
# 3) CCR summary shown in Excel sequence
# 4) CCR summary downloadable as CSV / Excel
# 5) Filled Import Declaration PDF generation
# =========================================================
# -----------------------------
# PAGE SETUP
# -----------------------------
st.set_page_config(page_title="GSO Expert Pro", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #F3F0F7; }
    [data-testid="stSidebar"] { background-color: #4B3F72 !important; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    h1, h2, h3 { color: #2E2841; font-family: 'Segoe UI', sans-serif; }
    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #D1C4E9;
        padding: 15px;
        border-radius: 12px;
    }
    .stButton>button {
        background: #7A61BA;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        border: none;
        height: 3em;
    }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #4B3F72;
        color: #FFFFFF;
        text-align: center;
        padding: 8px;
        z-index: 100;
        font-weight: bold;
    }
    </style>
    <div class="footer">MADE BY ABDULLAH ALHAKIM & ABDULLA DAABAL</div>
    """, unsafe_allow_html=True)
# -----------------------------
# FIREBASE SETUP
# -----------------------------
if not firebase_admin._apps:
    try:
        st.info("🔐 Attempting Firebase connection...")
        creds_dict = dict(st.secrets["firebase_credentials"])
        st.write("Credential fields found:", list(creds_dict.keys()))
        if "private_key" in creds_dict:
            private_key = creds_dict["private_key"]
            if "\\n" in private_key:
                private_key = private_key.replace("\\n", "\n")
                st.info("✓ Replaced escaped newlines")
            creds_dict["private_key"] = private_key
        cred = credentials.Certificate(creds_dict)
        firebase_admin.initialize_app(cred)
        st.success("✅ Firebase initialized successfully!")
    except Exception as e:
        st.error(f"❌ Database Connection Failed: {str(e)}")
        st.write("**Error Type:**", type(e).__name__)
        st.markdown("""
        ### 🔧 Troubleshooting Steps:
        1. Go to Firebase Console → Project Settings → Service Accounts
        2. Generate a NEW private key JSON
        3. Paste the FULL JSON into Streamlit secrets
        4. Keep private_key on one line using \\n
        """)
        st.stop()
# -----------------------------
# CLOUDINARY SETUP
# -----------------------------
try:
    cloudinary.config(
        cloud_name=st.secrets["cloudinary"]["cloud_name"],
        api_key=st.secrets["cloudinary"]["api_key"],
        api_secret=st.secrets["cloudinary"]["api_secret"],
        secure=True
    )
except Exception as e:
    st.error(f"❌ Cloudinary Configuration Failed: {e}")
    st.markdown("""
    ### 🔧 Cloudinary Secrets Needed:
    Add this in Streamlit Secrets:
    [cloudinary]
    cloud_name = "YOUR_CLOUD_NAME"
    api_key = "YOUR_API_KEY"
    api_secret = "YOUR_API_SECRET"
    """)
    st.stop()
db = firestore.client()
try:
    test_docs = list(db.collection("gso_database").limit(3).stream())
    st.write("Connected project:", st.secrets["firebase_credentials"]["project_id"])
    st.write("Documents found in gso_database:", len(test_docs))
except Exception as e:
    st.error(f"Firestore test failed: {e}")
# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def format_date_to_string(date_str):
    months = {
        "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
        "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
        "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"
    }
    try:
        parts = date_str.split()
        return f"{parts[0].zfill(2)}{months.get(parts[1].upper(), '00')}{parts[2][-2:]}"
    except Exception:
        return "000000"
def is_expired(expiry_ddmmyy):
    try:
        exp_date = datetime.strptime(expiry_ddmmyy, "%d%m%y")
        return exp_date.date() < datetime.today().date()
    except Exception:
        return True
def add_signature_to_pdf(page):
    text = " "
    page_rect = page.rect
    point = fitz.Point(page_rect.width - 200, page_rect.height - 20)
    page.insert_text(point, text, fontsize=10, color=(0.4, 0.4, 0.4))
def create_template(temp_type):
    output = io.BytesIO()
    if temp_type == "MICHELIN":
        df = pd.DataFrame(columns=["Ref Number", "Country"])
    else:
        df = pd.DataFrame(columns=["Brand", "Size", "Pattern"])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()
def create_ccr_summary_excel(ccr_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        ccr_df.to_excel(writer, index=False, sheet_name="CCR Summary")
    output.seek(0)
    return output.getvalue()
def extract_first_match(pattern, text, default=""):
    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
    return match.group(1).strip() if match else default
def normalize_pdf_text(text):
    """Make PDF text extraction more tolerant to spacing/OCR quirks."""
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*:\s*", ": ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()
def extract_field_by_label(text, label, default=""):
    pattern = rf"(?:^|\n){re.escape(label)}\s*:\s*(.+)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return default
def parse_expiry_date(text):
    match = re.search(r"Date of Expiry\s*:\s*(\d{1,2}\s*[A-Z]{3}\s*\d{4})", text, re.IGNORECASE)
    if match:
        return format_date_to_string(match.group(1).strip())
    return "000000"
def extract_certificate_fields(text):
    text = normalize_pdf_text(text)
    ccr_no = extract_first_match(r"CCR\s*No\.?\s*[:#-]?\s*([A-Z0-9]{5,})", text)
    brand = extract_field_by_label(text, "Brand").upper()
    pattern = extract_field_by_label(text, "Pattern").upper()
    country = extract_field_by_label(text, "Country of Production").upper()
    ref_no = extract_first_match(r"Manufacturer Ref No\s*:\s*([A-Z0-9-]+)", text).zfill(6)
    tyre_type = extract_field_by_label(text, "Type")
    expiry = parse_expiry_date(text)
    clean_size = tyre_type.replace("/", "-").strip().upper()
    return {
        "ccr_no": ccr_no,
        "brand": brand,
        "pattern": pattern,
        "country": country,
        "ref_no": ref_no,
        "size": clean_size,
        "expiry": expiry,
    }
def sanitize_doc_part(value):
    value = str(value or "").strip().upper()
    value = re.sub(r"[^A-Z0-9]+", "_", value)
    return value.strip("_")
def build_doc_id(fields):
    """
    Naming rule for uploaded certificate PDFs:
    BRAND_MANUFACTURERREFNO_COUNTRY_EXPIRY(DDMMYY)
    Example: MICHELIN_922041_FRANCE_240526
    """
    brand = sanitize_doc_part(fields.get("brand", ""))
    ref_no = sanitize_doc_part(fields.get("ref_no", ""))
    country = sanitize_doc_part(fields.get("country", ""))
    expiry = sanitize_doc_part(fields.get("expiry", ""))
    return f"{brand}_{ref_no}_{country}_{expiry}"
def get_import_decl_default_rect(page):
    """
    CCR box for the uploaded Import Declaration.
    Uses the full middle cell beside 'Conformity Certificate/s No:'.
    """
    w = page.rect.width
    h = page.rect.height
    return fitz.Rect(
        w * 0.394,   # left
        h * 0.324,   # top
        w * 0.663,   # right
        h * 0.417    # bottom
    )
def get_import_decl_safe_zones(page, allow_row_overflow=False):
    """
    Keep all CCRs strictly inside the main CCR box only.
    """
    return [get_import_decl_default_rect(page)]
def choose_multi_zone_layout(ccr_values, zones, requested_fontsize=None):
    count = len(ccr_values)
    if count <= 0:
        return {"font_size": 12, "line_height": 14, "zones": []}

    rect = fitz.Rect(zones[0].x0 + 6, zones[0].y0 + 6, zones[0].x1 - 6, zones[0].y1 - 6)

    if count <= 12:
        cols = 3
    elif count <= 36:
        cols = 4
    else:
        cols = 5   # 37 to 45

    rows = max(1, (count + cols - 1) // cols)

    if requested_fontsize is not None and float(requested_fontsize) > 0:
        font_size = float(requested_fontsize)
    else:
        max_font_by_width = rect.width / (cols * 3.25)
        max_font_by_height = rect.height / (rows * 1.08)
        font_size = min(24, max_font_by_width, max_font_by_height)
        font_size = max(11.5, font_size)

    line_height = font_size * 1.05

    return {
        "font_size": font_size,
        "line_height": line_height,
        "zones": [{
            "rect": rect,
            "rows": rows,
            "cols": cols,
            "capacity": rows * cols,
        }],
        "total_capacity": rows * cols,
    }
def draw_ccrs_across_safe_zones(page, ccr_values, zones, fontsize=None):
    layout = choose_multi_zone_layout(ccr_values, zones, requested_fontsize=fontsize)
    font_size = layout["font_size"]
    line_height = layout["line_height"]

    zone_info = layout["zones"][0]
    rect = zone_info["rect"]
    cols = zone_info["cols"]
    col_width = rect.width / cols

    # White inner writing area, keep borders visible
    page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)

    for idx, ccr in enumerate(ccr_values[:zone_info["capacity"]]):
        row = idx // cols
        col = idx % cols

        text_width = fitz.get_text_length(ccr, fontname="hebo", fontsize=font_size)
        x = rect.x0 + (col * col_width) + max((col_width - text_width) / 2, 0)
        y = rect.y0 + (row * line_height) + font_size + 1

        if y > rect.y1 - 2:
            break

        page.insert_text(
            fitz.Point(x, y),
            ccr,
            fontsize=font_size,
            fontname="hebo",
            color=(0, 0, 0),
            overlay=True,
        )
def fill_import_declaration_pdf(template_bytes, ccr_text,
                                x0=None, y0=None, x1=None, y1=None,
                                fontsize=None, allow_row_overflow=False):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]
    auto_rect = get_import_decl_default_rect(page)

    if None in (x0, y0, x1, y1):
        primary_rect = auto_rect
    else:
        primary_rect = fitz.Rect(x0, y0, x1, y1)

    ccr_values = [c.strip() for c in re.split(r"[,\n]+", str(ccr_text)) if c.strip()]
    zones = [primary_rect]

    if ccr_values:
        draw_ccrs_across_safe_zones(page, ccr_values, zones, fontsize=fontsize)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()
def render_pdf_first_page_to_png_bytes(pdf_bytes, zoom=1.35):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    page = doc[0]
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    return pix.tobytes('png')
def build_preview_ccr_values(count):
    base = 551520
    return [str(base + i).zfill(6) for i in range(max(0, int(count)))]
def get_report_based_preview_ccrs():
    report_results = st.session_state.get("report_results", {})
    found_ccrs = report_results.get("found_ccrs", [])
    ccr_list = [str(item.get("CCR No", "")).strip() for item in found_ccrs if str(item.get("CCR No", "")).strip()]
    return ccr_list
def upload_pdf_to_cloudinary(pdf_bytes, doc_id):
    """Upload through a temp file for better Cloudinary raw-PDF reliability."""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_bytes)
            tmp.flush()
            tmp_path = tmp.name
        upload_result = cloudinary.uploader.upload(
            tmp_path,
            resource_type="raw",
            folder="certificates",
            public_id=doc_id,
            format="pdf",
            overwrite=True,
            invalidate=True,
            unique_filename=False,
            use_filename=False,
        )
        secure_url = upload_result.get("secure_url") or upload_result.get("url")
        if not secure_url:
            raise ValueError("Cloudinary did not return a valid file URL")
        return secure_url
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
def download_pdf_from_url(pdf_url):
    
    last_error = None
    candidate_urls = [pdf_url]
    if "/image/upload/" in pdf_url:
        candidate_urls.append(pdf_url.replace("/image/upload/", "/raw/upload/"))
    if "/video/upload/" in pdf_url:
        candidate_urls.append(pdf_url.replace("/video/upload/", "/raw/upload/"))
    for url in list(dict.fromkeys(candidate_urls)):
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            return response.content
        except Exception as e:
            last_error = e
    raise last_error if last_error else ValueError("Unable to download PDF from Cloudinary URL")
def parse_legacy_doc_id(doc_id):
    """
    Legacy format:
    BRAND_ITEMCODE_COUNTRY_EXPIRY
    Example:
    MICHELIN_922041_FRANCE_240526
    """
    parts = str(doc_id).strip().split("_")
    if len(parts) < 4:
        return {}

    brand = parts[0].strip().upper()
    ref_no = parts[1].strip().upper().zfill(6)
    country = parts[2].strip().upper()
    expiry = parts[3].strip().upper()

    return {
        "brand": brand,
        "ref_no": ref_no,
        "country": country,
        "expiry": expiry,
    }


def migrate_legacy_firestore_records():
    repaired = 0
    skipped = 0
    failed = 0
    logs = []

    try:
        docs = db.collection("gso_database").get()
    except Exception as e:
        return 0, 0, 1, pd.DataFrame([{
            "doc_id": "collection_load",
            "status": f"Failed to load Firestore documents: {e}"
        }])

    for doc in docs:
        try:
            data = doc.to_dict() or {}
            doc_id = doc.id

            current_ccr = str(data.get("ccr_no", "")).strip()
            current_url = str(data.get("url", "")).strip()

            if current_ccr:
                skipped += 1
                continue

            legacy_parts = parse_legacy_doc_id(doc_id)

            if not current_url:
                failed += 1
                logs.append({
                    "doc_id": doc_id,
                    "status": "Failed: no PDF URL stored"
                })
                continue

            pdf_bytes = download_pdf_from_url(current_url)
            pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

            if len(pdf_doc) == 0:
                failed += 1
                logs.append({
                    "doc_id": doc_id,
                    "status": "Failed: empty PDF"
                })
                continue

            page1_text = normalize_pdf_text(pdf_doc[0].get_text())
            fields = extract_certificate_fields(page1_text)

            recovered_ccr = str(fields.get("ccr_no", "")).strip()

            if not recovered_ccr:
                failed += 1
                logs.append({
                    "doc_id": doc_id,
                    "status": "Failed: CCR not found in PDF"
                })
                continue

            doc.reference.set({
                "brand": fields.get("brand") or legacy_parts.get("brand", data.get("brand", "")),
                "pattern": fields.get("pattern", data.get("pattern", "")),
                "country": fields.get("country") or legacy_parts.get("country", data.get("country", "")),
                "ref_no": fields.get("ref_no") or legacy_parts.get("ref_no", data.get("ref_no", "")),
                "ccr_no": recovered_ccr,
                "size": fields.get("size", data.get("size", "")),
                "expiry": fields.get("expiry") or legacy_parts.get("expiry", data.get("expiry", "")),
                "url": current_url,
                "migrated_from_legacy": True,
                "updated_at": firestore.SERVER_TIMESTAMP
            }, merge=True)

            repaired += 1
            logs.append({
                "doc_id": doc_id,
                "status": f"Repaired | CCR {recovered_ccr}"
            })

        except Exception as e:
            failed += 1
            logs.append({
                "doc_id": doc.id if 'doc' in locals() else "unknown",
                "status": f"Failed: {e}"
            })

    return repaired, skipped, failed, pd.DataFrame(logs)              
# -----------------------------
# DATABASE LOADER
# -----------------------------
@st.cache_data(ttl=600)
def load_database_index():
    data = []
    try:
        docs_ref = db.collection("gso_database")
        docs = docs_ref.get()
        for doc in docs:
            d = doc.to_dict()
            d["id"] = doc.id
            data.append(d)
    except Exception as e:
        st.error(f"Error loading database: {e}")
        return pd.DataFrame(columns=[
            "brand", "size", "pattern", "ref_no", "ccr_no",
            "country", "expiry", "id", "url"
        ])
    df = pd.DataFrame(data)
    if df.empty:
        return pd.DataFrame(columns=[
            "brand", "size", "pattern", "ref_no", "ccr_no",
            "country", "expiry", "id", "url"
        ])
    cols = ["brand", "size", "pattern", "ref_no", "ccr_no", "country", "expiry"]
    for col in cols:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip().str.upper()
    if "url" not in df.columns:
        df["url"] = ""
    df["size"] = df["size"].str.replace("/", "-", regex=False)
    return df
# -----------------------------
# SIDEBAR
# -----------------------------
with st.sidebar:
    st.title("GSO Finder")
    menu = st.radio("WORKFLOW", ["Dashboard", "Add Certificates", "Search & Merge"])
# -----------------------------
# DASHBOARD
# -----------------------------
if menu == "Dashboard":
    st.title("📊 Control Center")
    today_display = datetime.now().strftime("%d %B %Y")
    if st.button("🔄 Refresh Database"):
        load_database_index.clear()
        st.success("Database cache refreshed!")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("System Date", today_display)
    with c2:
        st.metric("Database", "Online")
    st.markdown("### 📥 Templates")
    tc1, tc2 = st.columns(2)
    with tc1:
        st.download_button(
            "Download Michelin Template",
            create_template("MICHELIN"),
            "Michelin_Template.xlsx"
        )
    with tc2:
        st.download_button(
            "Download Others Template",
            create_template("OTHERS"),
            "Others_Template.xlsx"
        )
    st.markdown("### 🛠 Legacy Repair")

    if st.button("Repair Legacy Records"):
        with st.spinner("Repairing old Firestore records from stored PDFs..."):
            repaired, skipped, failed, log_df = migrate_legacy_firestore_records()
            load_database_index.clear()

        st.success(f"Done. Repaired: {repaired} | Skipped: {skipped} | Failed: {failed}")

        if not log_df.empty:
            st.subheader("Migration Log")
            st.dataframe(log_df, use_container_width=True)
                              
# -----------------------------
# ADD CERTIFICATES
# -----------------------------
elif menu == "Add Certificates":
    st.title("📥 Batch Upload")
    uploaded_pdfs = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)
    if st.button("Sync to Cloud"):
        if not uploaded_pdfs:
            st.warning("Please upload at least one PDF.")
            st.stop()
        progress_bar = st.progress(0)
        status_text = st.empty()
        upload_log = []
        for i, uploaded_file in enumerate(uploaded_pdfs):
            try:
                file_bytes = uploaded_file.read()
                doc = fitz.open(stream=file_bytes, filetype="pdf")
            except Exception as e:
                upload_log.append({
                    "file": uploaded_file.name,
                    "status": f"Failed to open PDF: {e}"
                })
                progress_bar.progress((i + 1) / len(uploaded_pdfs))
                continue
            processed_any_page = False
            if len(doc) == 0:
                upload_log.append({
                    "file": uploaded_file.name,
                    "status": "Skipped: empty PDF"
                })
                progress_bar.progress((i + 1) / len(uploaded_pdfs))
                continue
            page_num = 0
            try:
                text = doc[0].get_text()
                normalized_text = normalize_pdf_text(text)
                if "GSO Conformity Certificate" not in normalized_text:
                    upload_log.append({
                        "file": uploaded_file.name,
                        "status": "Skipped page 1: first page is not a GSO Conformity Certificate"
                    })
                else:
                    fields = extract_certificate_fields(normalized_text)
                    if not fields["ccr_no"]:
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": "Skipped page 1: CCR No not detected"
                        })
                    elif not fields["brand"] or not fields["ref_no"] or not fields["country"]:
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": f"Skipped page 1: missing Brand, Ref No, or Country | extracted={fields}"
                        })
                    elif is_expired(fields["expiry"]):
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": "Skipped page 1: certificate expired"
                        })
                    else:
                        doc_id = build_doc_id(fields)
                        # Upload only the first page of the file and discard all following pages.
                        new_doc = fitz.open()
                        new_doc.insert_pdf(doc, from_page=0, to_page=0)
                        pdf_url = upload_pdf_to_cloudinary(new_doc.tobytes(), doc_id)
                        db.collection("gso_database").document(doc_id).set({
                            "brand": fields["brand"],
                            "pattern": fields["pattern"],
                            "country": fields["country"],
                            "ref_no": fields["ref_no"],
                            "ccr_no": fields["ccr_no"],
                            "size": fields["size"],
                            "expiry": fields["expiry"],
                            "url": pdf_url,
                            "source_pages_saved": 1,
                            "naming_rule": "brand_refno_country_expiry_ddmmyy",
                            "updated_at": firestore.SERVER_TIMESTAMP
                        })
                        processed_any_page = True
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": f"Uploaded page 1 only | Saved as {doc_id}.pdf | CCR {fields['ccr_no']}"
                        })
            except Exception as e:
                upload_log.append({
                    "file": uploaded_file.name,
                    "status": f"Skipped page 1: {e}"
                })
            if not processed_any_page:
                upload_log.append({
                    "file": uploaded_file.name,
                    "status": "No valid certificate found on first page"
                })
            progress_bar.progress((i + 1) / len(uploaded_pdfs))
            status_text.text(f"Processed file {i + 1} of {len(uploaded_pdfs)}")
        load_database_index.clear()
        st.success("Sync Complete!")
        if upload_log:
            st.subheader("Upload Log")
            st.dataframe(pd.DataFrame(upload_log), use_container_width=True)
# -----------------------------
# SEARCH & MERGE
# -----------------------------
elif menu == "Search & Merge":
    st.title("🔍 Report Generation")
    mode = st.radio("Category", ["MICHELIN / BFG", "OTHER BRANDS"], horizontal=True)
    excel_file = st.file_uploader("Upload Excel", type=["xlsx"], key="excel_upload")
    import_decl_file = st.file_uploader(
        "Upload Import Declaration PDF (optional)",
        type=["pdf"],
        key="import_decl_upload"
    )
with st.expander("Import Declaration Placement Settings"):
    st.caption("Leave these blank to use auto-scaled coordinates for the uploaded template. Only override them if you want manual fine-tuning.")
    use_manual_decl_coords = st.checkbox("Use manual coordinates", value=False)
    if use_manual_decl_coords:
        st.warning("Manual values must match the actual PDF page size. For scanned PDFs, small values like 155 / 272 usually place the text in the wrong area.")
        decl_x0 = st.number_input("x0", value=772)
        decl_y0 = st.number_input("y0", value=944)
        decl_x1 = st.number_input("x1", value=1300)
        decl_y1 = st.number_input("y1", value=1216)
        decl_fontsize = st.number_input("Font Size (0 = auto)", value=0)
    else:
        decl_x0 = decl_y0 = decl_x1 = decl_y1 = decl_fontsize = None
    st.markdown("### 1️⃣ Generate Report")
    if excel_file and st.button("Generate Report"):
        with st.spinner("Loading Database Index..."):
            db_df = load_database_index()
        if db_df.empty:
            st.error("Database is empty or failed to load.")
            st.stop()
        try:
            df = pd.read_excel(excel_file).astype(str).apply(
                lambda x: x.str.replace(r"\.0$", "", regex=True).str.strip()
            )
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")
            st.stop()
        combined_pdf = fitz.open()
        missing = []
        found_ccrs = []
        progress_bar = st.progress(0)
        for index, row in df.iterrows():
            matches = pd.DataFrame()
            if mode == "MICHELIN / BFG":
                try:
                    t_ref = str(row.iloc[0]).strip().zfill(6)
                    t_country = str(row.iloc[1]).strip().upper()
                    matches = db_df[
                        (db_df["ref_no"] == t_ref) &
                        (db_df["country"] == t_country)
                    ]
                except Exception:
                    missing.append(f"Row {index + 2}: Invalid Excel structure for MICHELIN / BFG")
                    progress_bar.progress((index + 1) / len(df))
                    continue
            else:
                try:
                    t_brand = str(row.iloc[0]).strip().upper()
                    t_size = str(row.iloc[1]).strip().replace("/", "-").upper()
                    t_pattern = str(row.iloc[2]).strip().upper()
                    matches = db_df[
                        (db_df["brand"] == t_brand) &
                        (db_df["pattern"] == t_pattern) &
                        (db_df["size"] == t_size)
                    ]
                except Exception:
                    missing.append(f"Row {index + 2}: Invalid Excel structure for OTHER BRANDS")
                    progress_bar.progress((index + 1) / len(df))
                    continue
            if not matches.empty:
                found_item = matches.iloc[0]
                if is_expired(found_item["expiry"]):
                    missing.append(f"Row {index + 2}: Certificate Expired")
                else:
                    found_ccrs.append({
                        "Excel Row": index + 2,
                        "Brand": found_item.get("brand", ""),
                        "Pattern": found_item.get("pattern", ""),
                        "Size": found_item.get("size", ""),
                        "Manufacturer Ref No": found_item.get("ref_no", ""),
                        "CCR No": found_item.get("ccr_no", ""),
                        "Country": found_item.get("country", ""),
                        "Status": "Found"
                    })
                    try:
                        pdf_url = found_item["url"]
                        pdf_bytes = download_pdf_from_url(pdf_url)
                        match_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                        for page in match_doc:
                            add_signature_to_pdf(page)
                        combined_pdf.insert_pdf(match_doc)
                    except Exception as e:
                        missing.append(f"Row {index + 2}: Found in DB but PDF fetch failed ({e})")
            else:
                missing.append(f"Row {index + 2}: Not Found")
            progress_bar.progress((index + 1) / len(df))
        report_pdf_bytes = None
        if len(combined_pdf) > 0:
            out = io.BytesIO()
            combined_pdf.save(out)
            out.seek(0)
            report_pdf_bytes = out.getvalue()
        st.session_state["report_results"] = {
            "found_ccrs": found_ccrs,
            "missing": missing,
            "report_pdf_bytes": report_pdf_bytes,
            "mode": mode,
        }
        default_editor_text = ", ".join([str(item["CCR No"]).strip() for item in found_ccrs if str(item["CCR No"]).strip()])
        st.session_state["final_decl_ccr_editor"] = default_editor_text
        st.session_state["custom_ccrs_text"] = default_editor_text
        st.success("Report generated. You can now review the CCR list and use it for live preview or final declaration.")
    report_results = st.session_state.get("report_results", {})
    found_ccrs = report_results.get("found_ccrs", [])
    missing = report_results.get("missing", [])
    report_pdf_bytes = report_results.get("report_pdf_bytes")
    if report_pdf_bytes:
        st.success(f"Generated {len(found_ccrs)} matched certificate page(s) in the report")
        st.download_button(
            "📥 DOWNLOAD REPORT",
            report_pdf_bytes,
            "GSO_Final_Report.pdf",
            "application/pdf"
        )
    if found_ccrs:
        ccr_df = pd.DataFrame(found_ccrs)
        st.subheader("Matched CCR Numbers in Excel Sequence")
        st.dataframe(ccr_df, use_container_width=True)
        csv_data = ccr_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download CCR Summary CSV",
            data=csv_data,
            file_name="CCR_Summary.csv",
            mime="text/csv"
        )
        excel_data = create_ccr_summary_excel(ccr_df)
        st.download_button(
            "Download CCR Summary Excel",
            data=excel_data,
            file_name="CCR_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    if missing:
        with st.expander("Errors / Missing"):
            for m in missing:
                st.error(m)
    if import_decl_file is not None:
        st.markdown("### 2️⃣ Live Preview")
        report_based_ccrs = get_report_based_preview_ccrs()
        if not report_based_ccrs:
            st.info("Generate the report first. The live preview custom CCR list now starts from the CCRs found in the report.")
        else:
            preview_mode = st.radio(
                "Preview source",
                ["Report CCR count", "Custom CCR list"],
                horizontal=True,
                key="preview_mode"
            )
            allow_row_overflow = st.checkbox(
                "Allow row overflow for large batches",
                value=False,
                help="For large CCR counts, continue across the same safe row without covering label text.",
                key="allow_row_overflow_preview"
            )
            if preview_mode == "Report CCR count":
                max_count = max(1, len(report_based_ccrs))
                preview_count = st.slider(
                    "Preview CCR count from report",
                    min_value=1,
                    max_value=max_count,
                    value=max_count,
                    key="preview_count"
                )
                preview_ccrs = report_based_ccrs[:preview_count]
            else:
                default_custom_text = st.session_state.get("custom_ccrs_text", ", ".join(report_based_ccrs))
                custom_ccrs_text = st.text_area(
                    "Custom CCRs based on report results (comma or line separated)",
                    value=default_custom_text,
                    key="custom_ccrs_text"
                )
                preview_ccrs = [c.strip() for c in re.split(r"[,\n]+", custom_ccrs_text) if c.strip()]
                preview_count = len(preview_ccrs)
            st.caption(f"Previewing {preview_count} CCR(s) from the report-based list")
            if st.button("Generate Live Preview", key="generate_live_preview"):
                try:
                    template_bytes = import_decl_file.getvalue()
                    preview_fontsize = None if decl_fontsize in (None, 0) else decl_fontsize
                    preview_pdf = fill_import_declaration_pdf(
                        template_bytes=template_bytes,
                        ccr_text=", ".join(preview_ccrs),
                        x0=decl_x0, y0=decl_y0, x1=decl_x1, y1=decl_y1,
                        fontsize=preview_fontsize,
                        allow_row_overflow=allow_row_overflow,
                    )
                    preview_png = render_pdf_first_page_to_png_bytes(preview_pdf)
                    st.image(preview_png, caption="Live preview of the filled Import Declaration", use_container_width=True)
                    st.download_button(
                        "Download Preview PDF",
                        data=preview_pdf,
                        file_name="Import_Declaration_Preview.pdf",
                        mime="application/pdf",
                        key="download_preview_pdf"
                    )
                except Exception as e:
                    st.error(f"Could not generate live preview: {e}")
    if found_ccrs and import_decl_file is not None:
        st.markdown("### 3️⃣ Final CCR List for Declaration")
        auto_ccr_list = [str(item["CCR No"]).strip() for item in found_ccrs if str(item["CCR No"]).strip()]
        auto_ccr_text = ", ".join(auto_ccr_list)
        st.caption("This list is auto-generated from the Excel matches below. You can edit it before exporting the declaration.")
        editable_ccr_text = st.text_area(
            "Edit CCRs for final declaration (comma or line separated)",
            value=st.session_state.get("final_decl_ccr_editor", auto_ccr_text),
            height=120,
            key="final_decl_ccr_editor"
        )
        final_ccr_list = [c.strip() for c in re.split(r"[,\n]+", editable_ccr_text) if c.strip()]
        final_ccr_text = ", ".join(final_ccr_list)
        if final_ccr_text:
            st.info(f"Final declaration will use {len(final_ccr_list)} CCR(s): {final_ccr_text}")
            try:
                filled_pdf = fill_import_declaration_pdf(
                    template_bytes=import_decl_file.getvalue(),
                    ccr_text=final_ccr_text,
                    x0=decl_x0, y0=decl_y0, x1=decl_x1, y1=decl_y1,
                    fontsize=(None if decl_fontsize in (None, 0) else decl_fontsize),
                    allow_row_overflow=False,
                )
                st.success("Import Declaration is ready using the edited CCR list.")
                st.download_button(
                    "Download Filled Import Declaration",
                    data=filled_pdf,
                    file_name="Filled_Import_Declaration.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Could not generate filled Import Declaration PDF: {e}")
        else:
            st.warning("No CCR numbers are currently entered for the declaration.")
