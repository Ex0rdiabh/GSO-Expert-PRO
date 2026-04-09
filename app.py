import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import fitz
import pandas as pd
import io
import re
import requests
import cloudinary
import cloudinary.uploader
from datetime import datetime
import math

# -----------------------------
# CONFIG
# -----------------------------
st.set_page_config(page_title="GSO Expert Pro", layout="wide")

# -----------------------------
# FIREBASE
# -----------------------------
if not firebase_admin._apps:
    cred = credentials.Certificate(dict(st.secrets["firebase_credentials"]))
    firebase_admin.initialize_app(cred)

db = firestore.client()

# -----------------------------
# CLOUDINARY
# -----------------------------
cloudinary.config(
    cloud_name=st.secrets["cloudinary"]["cloud_name"],
    api_key=st.secrets["cloudinary"]["api_key"],
    api_secret=st.secrets["cloudinary"]["api_secret"],
    secure=True
)

# -----------------------------
# HELPERS
# -----------------------------
def download_pdf_from_url(url):
    r = requests.get(url)
    r.raise_for_status()
    return r.content

def upload_pdf_to_cloudinary(pdf_bytes, doc_id):
    result = cloudinary.uploader.upload(
        pdf_bytes,
        resource_type="raw",
        public_id=doc_id,
        format="pdf",
        overwrite=True
    )
    return result["secure_url"]

def extract_first(pattern, text):
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def extract_fields(text):
    return {
        "ccr_no": extract_first(r"CCR No:\s*(\d+)", text),
        "ref_no": extract_first(r"Manufacturer Ref No:\s*(.*)", text).zfill(6),
        "country": extract_first(r"Country of Production:\s*(.*)", text).upper()
    }

# -----------------------------
# CCR GRID LAYOUT (SMART)
# -----------------------------
def draw_ccr_grid(page, ccr_list, rect):
    x0, y0, x1, y1 = rect
    width = x1 - x0
    height = y1 - y0

    n = len(ccr_list)

    # columns logic
    if n <= 4:
        cols = 2
    elif n <= 10:
        cols = 3
    elif n <= 20:
        cols = 4
    elif n <= 30:
        cols = 5
    else:
        cols = 6

    rows = math.ceil(n / cols)

    col_w = width / cols
    row_h = height / rows

    font = max(10, min(28, int(row_h * 0.6)))

    for i, ccr in enumerate(ccr_list):
        r = i // cols
        c = i % cols

        x = x0 + c * col_w + 5
        y = y0 + r * row_h + 15

        page.insert_text((x, y), ccr, fontsize=font)

# -----------------------------
# FILL DECLARATION
# -----------------------------
def fill_pdf(template, ccr_list):
    doc = fitz.open(stream=template, filetype="pdf")
    page = doc[0]

    # MAIN BOX
    rect = (775, 1045, 1325, 1170)

    draw_ccr_grid(page, ccr_list, rect)

    # OVERFLOW ROW
    if len(ccr_list) > 30:
        rect2 = (1325, 1045, 1800, 1170)
        draw_ccr_grid(page, ccr_list[30:], rect2)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -----------------------------
# UI
# -----------------------------
st.title("GSO Expert Pro")

menu = st.sidebar.radio("Menu", ["Add Certificates", "Search & Merge"])

# -----------------------------
# ADD CERTIFICATES
# -----------------------------
if menu == "Add Certificates":
    files = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)

    if st.button("Upload"):
        for file in files:
            doc = fitz.open(stream=file.read(), filetype="pdf")

            for page in doc:
                text = page.get_text()

                if "GSO Conformity Certificate" not in text:
                    continue

                f = extract_fields(text)

                if not f["ref_no"]:
                    st.warning("Skipped page")
                    continue

                doc_id = f"{f['ref_no']}_{f['country']}"

                new_doc = fitz.open()
                new_doc.insert_pdf(doc)

                url = upload_pdf_to_cloudinary(new_doc.tobytes(), doc_id)

                db.collection("gso_database").document(doc_id).set({
                    **f,
                    "url": url
                })

                st.success(f"Uploaded {f['ccr_no']}")

# -----------------------------
# SEARCH
# -----------------------------
if menu == "Search & Merge":

    excel = st.file_uploader("Upload Excel", type="xlsx")
    decl = st.file_uploader("Upload Declaration PDF", type="pdf")

    if excel and st.button("Generate Report"):
        df = pd.read_excel(excel)

        docs = db.collection("gso_database").get()

        db_df = pd.DataFrame([d.to_dict() for d in docs])

        combined = fitz.open()
        found = []

        for _, row in df.iterrows():
            ref = str(row.iloc[0]).zfill(6)
            country = str(row.iloc[1]).upper()

            match = db_df[
                (db_df["ref_no"] == ref) &
                (db_df["country"] == country)
            ]

            if not match.empty:
                item = match.iloc[0]
                found.append(item["ccr_no"])

                pdf = download_pdf_from_url(item["url"])
                doc = fitz.open(stream=pdf, filetype="pdf")
                combined.insert_pdf(doc)

        # DOWNLOAD REPORT
        if len(combined) > 0:
            out = io.BytesIO()
            combined.save(out)

            st.download_button(
                "Download Report",
                out.getvalue(),
                "report.pdf"
            )

        # DECLARATION
        if decl and found:
            filled = fill_pdf(decl.read(), found)

            st.download_button(
                "Download Declaration",
                filled,
                "declaration.pdf"
            )

    # -----------------------------
    # LIVE PREVIEW
    # -----------------------------
    st.subheader("Live Preview")

    preview_input = st.text_area("Enter CCRs (comma or new line)")

    if st.button("Preview"):
        preview_ccrs = [
            c.strip()
            for c in re.split(r"[,\n]+", preview_input)
            if c.strip()
        ]

        if decl and preview_ccrs:
            preview_pdf = fill_pdf(decl.read(), preview_ccrs)

            st.download_button(
                "Download Preview",
                preview_pdf,
                "preview.pdf"
            )
