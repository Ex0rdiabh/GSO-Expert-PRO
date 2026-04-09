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

    ccr_no = extract_first_match(r"CCR No\s*:\s*(\d{5,})", text)
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


def build_doc_id(fields):
    brand = fields["brand"]
    ref_no = fields["ref_no"]
    country = fields["country"]
    expiry = fields["expiry"]
    size = fields["size"]
    pattern = fields["pattern"]

    if brand in ["MICHELIN", "BFGOODRICH"]:
        return f"{brand}_{ref_no}_{country}_{expiry}"
    return f"{brand}_{size}_{pattern}_{expiry}"


def insert_wrapped_text(page, text, rect, fontsize=10, align=0):
    page.insert_textbox(
        rect,
        text,
        fontsize=fontsize,
        fontname="helv",
        color=(0, 0, 0),
        align=align
    )


def fill_import_declaration_pdf(template_bytes, ccr_text,
                                x0=155, y0=272, x1=505, y1=322,
                                fontsize=11):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]

    rect = fitz.Rect(x0, y0, x1, y1)
    insert_wrapped_text(page, ccr_text, rect, fontsize=fontsize, align=0)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


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

            for page_num in range(len(doc)):
                try:
                    text = doc[page_num].get_text()
                    normalized_text = normalize_pdf_text(text)

                    if "GSO Conformity Certificate" not in normalized_text:
                        continue

                    fields = extract_certificate_fields(normalized_text)

                    if not fields["ccr_no"]:
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": f"Skipped page {page_num + 1}: CCR No not detected"
                        })
                        continue

                    if not fields["ref_no"] or not fields["country"]:
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": f"Skipped page {page_num + 1}: missing Ref No or Country | extracted={fields}"
                        })
                        continue

                    if is_expired(fields["expiry"]):
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": f"Skipped page {page_num + 1}: certificate expired"
                        })
                        continue

                    doc_id = build_doc_id(fields)

                    new_doc = fitz.open()
                    end_page = page_num
                    if page_num + 1 < len(doc):
                        next_text = normalize_pdf_text(doc[page_num + 1].get_text())
                        if "Passenger Car Tyres Test Report" in next_text and f"CCR No: {fields['ccr_no']}" in next_text:
                            end_page = page_num + 1
                    new_doc.insert_pdf(doc, from_page=page_num, to_page=end_page)

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
                        "updated_at": firestore.SERVER_TIMESTAMP
                    })

                    processed_any_page = True
                    upload_log.append({
                        "file": uploaded_file.name,
                        "status": f"Uploaded | CCR {fields['ccr_no']} | Ref {fields['ref_no']}"
                    })

                except Exception as e:
                    upload_log.append({
                        "file": uploaded_file.name,
                        "status": f"Skipped page {page_num + 1}: {e}"
                    })

            if not processed_any_page:
                upload_log.append({
                    "file": uploaded_file.name,
                    "status": "No valid certificate pages found"
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
            decl_x0 = st.number_input("x0", value=775)
            decl_y0 = st.number_input("y0", value=1045)
            decl_x1 = st.number_input("x1", value=1325)
            decl_y1 = st.number_input("y1", value=1170)
            decl_fontsize = st.number_input("Font Size (0 = auto)", value=0)
        else:
            decl_x0 = decl_y0 = decl_x1 = decl_y1 = decl_fontsize = None

    if import_decl_file is not None:
        st.markdown("### 👁️ Live Preview")
        preview_mode = st.radio(
            "Preview source",
            ["Sample CCR count", "Custom CCR list"],
            horizontal=True,
            key="preview_mode"
        )
        allow_row_overflow = st.checkbox(
            "Allow row overflow for large batches",
            value=True,
            help="For large CCR counts, continue across the same safe row without covering label text.",
            key="allow_row_overflow_preview"
        )

        if preview_mode == "Sample CCR count":
            preview_count = st.slider("Preview CCR count", min_value=1, max_value=40, value=20, key="preview_count")
            preview_ccrs = build_preview_ccr_values(preview_count)
        else:
            custom_ccrs_text = st.text_area(
                "Custom CCRs (comma or line separated)",
                value="551520, 561750, 568580, 570876",
                key="custom_ccrs_text"
            )
            preview_ccrs = [c.strip() for c in re.split(r"[,\n]+", custom_ccrs_text) if c.strip()]
            preview_count = len(preview_ccrs)

        st.caption(f"Previewing {preview_count} CCR(s)")

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

        if len(combined_pdf) > 0:
            out = io.BytesIO()
            combined_pdf.save(out)
            out.seek(0)
            st.success(f"Generated {len(combined_pdf)} Page PDF")
            st.download_button(
                "📥 DOWNLOAD REPORT",
                out.getvalue(),
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

            if import_decl_file is not None:
                ccr_list = [str(item["CCR No"]).strip() for item in found_ccrs if str(item["CCR No"]).strip()]
                ccr_text = ", ".join(ccr_list)

                if ccr_text:
                    try:
                        filled_pdf = fill_import_declaration_pdf(
                            template_bytes=import_decl_file.read(),
                            ccr_text=ccr_text,
                            x0=decl_x0, y0=decl_y0, x1=decl_x1, y1=decl_y1,
                            fontsize=(None if decl_fontsize == 0 else decl_fontsize)
                        )
                        st.success(f"Import Declaration filled with CCR text: {ccr_text}")

                        st.download_button(
                            "Download Filled Import Declaration",
                            data=filled_pdf,
                            file_name="Filled_Import_Declaration.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"Could not generate filled Import Declaration PDF: {e}")
                else:
                    st.warning("No CCR numbers were found to place into the Import Declaration PDF.")

        if missing:
            with st.expander("Errors / Missing"):
                for m in missing:
                    st.error(m)
