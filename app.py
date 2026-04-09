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
from datetime import datetime
import math

# =========================================================
# GSO FINDER - CLOUDINARY + FIRESTORE VERSION (PATCHED)
# Preserves original UI/dashboard/footer and adds:
# 1) Stable Cloudinary upload/download helpers
# 2) Improved CCR placement for Import Declaration
# 3) Live preview for Import Declaration before download
# 4) Top-left aligned CCR layout with same-row overflow for large batches
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
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else default


def extract_certificate_fields(text):
    ccr_no = extract_first_match(r"CCR No:\s*(\d+)", text)
    brand = extract_first_match(r"Brand:\s*(.*)", text).upper()
    pattern = extract_first_match(r"Pattern:\s*(.*)", text).upper()
    country = extract_first_match(r"Country of Production:\s*(.*)", text).upper()
    ref_no = extract_first_match(r"Manufacturer Ref No:\s*(.*)", text).zfill(6)
    tyre_type = extract_first_match(r"Type:\s*(.*)", text)
    expiry_raw = extract_first_match(
        r"Date of Expiry:\s*(\d{1,2}\s*[A-Z]{3}\s*\d{4})",
        text
    )
    expiry = format_date_to_string(expiry_raw) if expiry_raw else "000000"
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


def upload_pdf_to_cloudinary(pdf_bytes, doc_id):
    upload_result = cloudinary.uploader.upload(
        pdf_bytes,
        resource_type="raw",
        folder="certificates",
        public_id=doc_id,
        format="pdf",
        overwrite=True
    )
    return upload_result["secure_url"]


def download_pdf_from_url(pdf_url):
    response = requests.get(pdf_url, timeout=30)
    response.raise_for_status()
    return response.content


def _choose_layout(n):
    if n <= 4:
        return 2
    if n <= 8:
        return 3
    if n <= 15:
        return 4
    if n <= 24:
        return 5
    return 6


def _fit_font_for_grid(rect, rows, cols, max_font=30, min_font=8):
    width = max(rect.width, 1)
    height = max(rect.height, 1)
    usable_w = max(width - 14, 1)
    usable_h = max(height - 12, 1)

    font_by_height = usable_h / max(rows, 1) * 0.48
    font_by_width = usable_w / max(cols, 1) / 3.9
    font = min(max_font, font_by_height, font_by_width)
    return max(min_font, font)


def _grid_lines(ccr_list, cols):
    lines = []
    for i in range(0, len(ccr_list), cols):
        lines.append(ccr_list[i:i + cols])
    return lines


def _draw_ccrs_in_rect(page, ccr_list, rect, cols, font_size=0, top_left=True):
    if not ccr_list:
        return

    lines = _grid_lines(ccr_list, cols)
    rows = len(lines)

    inner_left = rect.x0 + 8
    inner_top = rect.y0 + 8
    inner_right = rect.x1 - 8
    inner_bottom = rect.y1 - 8

    inner_rect = fitz.Rect(inner_left, inner_top, inner_right, inner_bottom)

    if font_size and font_size > 0:
        font = float(font_size)
    else:
        font = _fit_font_for_grid(inner_rect, rows, cols)

    cell_w = max((inner_rect.width / cols), 1)
    cell_h = max((inner_rect.height / rows), 1)
    baseline_offset = max(font * 0.95, 8)

    for r, row_items in enumerate(lines):
        for c, ccr in enumerate(row_items):
            x = inner_rect.x0 + c * cell_w
            y = inner_rect.y0 + r * cell_h
            cell_rect = fitz.Rect(x, y, x + cell_w, y + cell_h)
            draw_x = cell_rect.x0 if top_left else cell_rect.x0 + 2
            draw_y = min(cell_rect.y0 + baseline_offset, cell_rect.y1 - 2)
            page.insert_text(
                fitz.Point(draw_x, draw_y),
                str(ccr),
                fontsize=font,
                fontname="helv",
                color=(0, 0, 0)
            )


def fill_import_declaration_pdf(
    template_bytes,
    ccr_text,
    x0=775,
    y0=1045,
    x1=1325,
    y1=1170,
    fontsize=0,
    overflow_enabled=True,
):
    """
    Place CCR numbers from the top-left of the main CCR box.
    For large batches, continue across the same row to the right,
    while staying above the row's lower text region.
    """
    ccr_list = [c.strip() for c in re.split(r"[,\n]+", ccr_text) if c.strip()]
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]

    main_rect = fitz.Rect(float(x0), float(y0), float(x1), float(y1))

    # draw inside main middle box first
    if len(ccr_list) <= 24 or not overflow_enabled:
        cols = _choose_layout(len(ccr_list))
        _draw_ccrs_in_rect(page, ccr_list, main_rect, cols, fontsize, top_left=True)
    else:
        # Keep first 24 in the main box, then continue in right-side safe zone on same row.
        main_count = 24
        first_batch = ccr_list[:main_count]
        remaining = ccr_list[main_count:]

        main_cols = _choose_layout(len(first_batch))
        _draw_ccrs_in_rect(page, first_batch, main_rect, main_cols, fontsize, top_left=True)

        safe_right_margin = 80
        gap = 18
        overflow_x0 = min(main_rect.x1 + gap, page.rect.width - safe_right_margin - 40)
        overflow_x1 = max(overflow_x0 + 40, page.rect.width - safe_right_margin)
        overflow_rect = fitz.Rect(overflow_x0, main_rect.y0, overflow_x1, main_rect.y1)

        overflow_cols = _choose_layout(len(remaining))
        _draw_ccrs_in_rect(page, remaining, overflow_rect, overflow_cols, fontsize, top_left=True)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


def render_pdf_preview_image(pdf_bytes, zoom=1.15):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return pix.tobytes("png")


# -----------------------------
# DATABASE LOADER
# -----------------------------
@st.cache_data(ttl=600)
def load_database_index():
    data = []

    try:
        docs_ref = db.collection("gso_database")
        docs = docs_ref.get()

        for doc_item in docs:
            d = doc_item.to_dict()
            d["id"] = doc_item.id
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

            for page_num in range(0, len(doc), 2):
                try:
                    text = doc[page_num].get_text()

                    if "GSO Conformity Certificate" not in text:
                        continue

                    fields = extract_certificate_fields(text)

                    if not fields["ref_no"] or not fields["country"]:
                        upload_log.append({
                            "file": uploaded_file.name,
                            "status": f"Skipped page {page_num + 1}: missing Ref No or Country"
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
                    end_page = min(page_num + 1, len(doc) - 1)
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
        st.caption("Adjust only if the CCR text is not aligned correctly in the declaration PDF.")
        decl_x0 = st.number_input("x0", value=775)
        decl_y0 = st.number_input("y0", value=1045)
        decl_x1 = st.number_input("x1", value=1325)
        decl_y1 = st.number_input("y1", value=1170)
        decl_fontsize = st.number_input("Font Size (0 = auto)", value=0)
        decl_overflow = st.checkbox("Allow same-row overflow for large batches", value=True)

    import_decl_bytes = import_decl_file.getvalue() if import_decl_file is not None else None

    # -------- Live Preview --------
    if import_decl_bytes is not None:
        st.markdown("### 👁️ Live Preview")
        preview_mode = st.radio(
            "Preview Input",
            ["Sample Count", "Custom CCR List"],
            horizontal=True,
            key="preview_mode"
        )

        if preview_mode == "Sample Count":
            preview_count = st.slider("Number of CCRs to preview", min_value=1, max_value=40, value=12)
            preview_ccrs = [str(551520 + i) for i in range(preview_count)]
        else:
            preview_input = st.text_area(
                "Paste CCRs separated by commas or new lines",
                value="551520, 561750, 568580, 570876",
                height=120
            )
            preview_ccrs = [c.strip() for c in re.split(r"[,\n]+", preview_input) if c.strip()]

        if st.button("Generate Live Preview"):
            if preview_ccrs:
                preview_pdf = fill_import_declaration_pdf(
                    template_bytes=import_decl_bytes,
                    ccr_text=", ".join(preview_ccrs),
                    x0=decl_x0, y0=decl_y0, x1=decl_x1, y1=decl_y1,
                    fontsize=decl_fontsize,
                    overflow_enabled=decl_overflow,
                )
                preview_png = render_pdf_preview_image(preview_pdf)
                st.image(preview_png, caption="Import Declaration Preview", use_container_width=True)
                st.download_button(
                    "Download Preview PDF",
                    data=preview_pdf,
                    file_name="Preview_Import_Declaration.pdf",
                    mime="application/pdf"
                )
            else:
                st.warning("Please add at least one CCR number for preview.")

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

            if import_decl_bytes is not None:
                ccr_list = [str(item["CCR No"]).strip() for item in found_ccrs if str(item["CCR No"]).strip()]
                ccr_text = ", ".join(ccr_list)

                if ccr_text:
                    try:
                        filled_pdf = fill_import_declaration_pdf(
                            template_bytes=import_decl_bytes,
                            ccr_text=ccr_text,
                            x0=decl_x0, y0=decl_y0, x1=decl_x1, y1=decl_y1,
                            fontsize=decl_fontsize,
                            overflow_enabled=decl_overflow,
                        )

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
