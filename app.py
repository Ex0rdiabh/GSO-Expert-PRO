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


def clean_ccr_list(ccr_text):
    return [v.strip() for v in str(ccr_text).split(",") if v.strip()]


def get_import_decl_default_rect(page):
    """
    Tuned for the uploaded scanned declaration template.
    This rectangle targets the middle box for:
    'Conformity Certificate/s No:'.
    """
    return fitz.Rect(775, 1045, 1325, 1170)


def get_import_decl_safe_zones(page, allow_row_overflow=False):
    """
    Safe writable areas for the CCR row.

    Normal mode:
        - middle box only
    Large-batch overflow mode:
        - middle box first
        - then the blank area of the right box, without covering the Arabic label
        - then the blank area of the left box, without covering the English label
    """
    middle = fitz.Rect(775, 1045, 1325, 1170)

    if not allow_row_overflow:
        return [middle]

    # Tuned for the uploaded declaration scan rendered in the app.
    # Right blank area: keep clear of the Arabic title on the top-right.
    right_blank = fitz.Rect(1335, 1045, 1710, 1170)

    # Left blank area: start after the English label text so we do not cover it.
    left_blank = fitz.Rect(505, 1045, 760, 1170)

    return [middle, right_blank, left_blank]


def choose_multi_zone_layout(ccr_values, zones, requested_fontsize=None):
    """
    Choose a font size and column count per safe zone so all CCRs fit.
    For large batches, allows overflow across the same row while avoiding label text.
    """
    count = len(ccr_values)
    if count <= 0:
        return {
            "font_size": 12,
            "line_height": 14,
            "zones": []
        }

    if requested_fontsize is not None and float(requested_fontsize) > 0:
        font_candidates = [float(requested_fontsize)]
    else:
        # Search from larger to smaller fonts until the total capacity fits the CCR count.
        font_candidates = [x / 10 for x in range(28 * 10, 7 * 10 - 1, -2)]

    best = None

    for font_size in font_candidates:
        line_height = max(font_size * 1.12, font_size + 2)
        zone_layouts = []
        total_capacity = 0

        for zi, zone in enumerate(zones):
            inner = fitz.Rect(zone.x0 + 8, zone.y0 + 8, zone.x1 - 8, zone.y1 - 8)
            rows = max(1, int(inner.height // line_height))

            # Slightly denser packing in the central zone.
            char_factor = 4.05 if zi == 0 else 4.15
            col_width_needed = font_size * char_factor
            cols = max(1, int(inner.width // col_width_needed))

            capacity = rows * cols
            total_capacity += capacity
            zone_layouts.append({
                'rect': inner,
                'rows': rows,
                'cols': cols,
                'capacity': capacity,
            })

        best = {
            'font_size': font_size,
            'line_height': line_height,
            'zones': zone_layouts,
            'total_capacity': total_capacity,
        }

        if total_capacity >= count:
            return best

    return best


def draw_ccrs_across_safe_zones(page, ccr_values, zones, fontsize=None):
    """
    Draw CCRs from the top-left of the middle box first.
    For large batches, continue into adjacent safe areas on the same row,
    while staying inside those areas and not covering the labels.
    """
    layout = choose_multi_zone_layout(ccr_values, zones, requested_fontsize=fontsize)
    font_size = layout['font_size']
    line_height = layout['line_height']

    remaining = list(ccr_values)

    # White only the writable zones, not the labels.
    for zone_info in layout['zones']:
        page.draw_rect(zone_info['rect'], color=(1, 1, 1), fill=(1, 1, 1), overlay=True)

    for zone_info in layout['zones']:
        if not remaining:
            break

        rect = zone_info['rect']
        rows = zone_info['rows']
        cols = zone_info['cols']
        col_width = rect.width / cols
        start_x = rect.x0 + 2
        start_y = rect.y0 + font_size

        for idx_in_zone in range(min(zone_info['capacity'], len(remaining))):
            ccr = remaining.pop(0)
            row = idx_in_zone // cols
            col = idx_in_zone % cols
            x = start_x + (col * col_width)
            y = start_y + (row * line_height)

            if y > rect.y1 - 1:
                break

            page.insert_text(
                fitz.Point(x, y),
                ccr,
                fontsize=font_size,
                fontname='cour',
                color=(0, 0, 0),
                overlay=True,
            )


def fill_import_declaration_pdf(template_bytes, ccr_text,
                                x0=None, y0=None, x1=None, y1=None,
                                fontsize=None, allow_row_overflow=True):
    doc = fitz.open(stream=template_bytes, filetype='pdf')
    page = doc[0]

    auto_rect = get_import_decl_default_rect(page)

    if None in (x0, y0, x1, y1):
        primary_rect = auto_rect
    else:
        primary_rect = fitz.Rect(x0, y0, x1, y1)

    ccr_values = [c.strip() for c in str(ccr_text).split(',') if c.strip()]

    if allow_row_overflow and len(ccr_values) > 18 and None in (x0, y0, x1, y1):
        zones = get_import_decl_safe_zones(page, allow_row_overflow=True)
    else:
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


def generate_import_decl_preview(template_bytes, count=20, fontsize=None, allow_row_overflow=True,
                                 x0=None, y0=None, x1=None, y1=None):
    ccr_values = build_preview_ccr_values(count)
    ccr_text = ', '.join(ccr_values)
    filled_pdf = fill_import_declaration_pdf(
        template_bytes=template_bytes,
        ccr_text=ccr_text,
        x0=x0, y0=y0, x1=x1, y1=y1,
        fontsize=fontsize,
        allow_row_overflow=allow_row_overflow,
    )
    return filled_pdf, render_pdf_first_page_to_png_bytes(filled_pdf)

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
