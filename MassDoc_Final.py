import streamlit as st
import os
import fitz  # PyMuPDF
from PIL import Image, ImageDraw
import zipfile
import pandas as pd
from pdf2docx import Converter
from docx2pdf import convert
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# ================= CONFIG =================
st.set_page_config(page_title="MassDoc Converter", layout="centered")

st.title("🧰 MassDoc Converter")
st.caption("Upload → Convert → Download")

st.divider()
school_mode = st.toggle("🏫 Aktifkan Mode Sekolah")
st.divider()

# ================= FUNCTIONS =================
def save_temp(uploaded_file):
    path = f"temp_{uploaded_file.name}"
    with open(path, "wb") as f:
        f.write(uploaded_file.read())
    return path

def pdf_to_png(pdf_path, output_dir, dpi):
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    doc = fitz.open(pdf_path)
    images = []

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        out = os.path.join(output_dir, f"page_{i+1}.png")
        pix.save(out)
        images.append(out)
    return images

def pdf_to_word(pdf_path, output_docx):
    cv = Converter(pdf_path)
    cv.convert(output_docx)
    cv.close()

def word_to_pdf(docx_path, output_pdf):
    convert(docx_path, output_pdf)

def excel_to_pdf(excel_path, pdf_path):
    df = pd.read_excel(excel_path)
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4
    x_start, y = 40, height - 40

    for _, row in df.iterrows():
        x = x_start
        for cell in row:
            c.drawString(x, y, str(cell))
            x += 100
        y -= 20
        if y < 40:
            c.showPage()
            y = height - 40
    c.save()

def add_watermark(img_path, text):
    img = Image.open(img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    draw.text((20, img.height - 40), text, fill=(150, 150, 150, 120))
    img.save(img_path)

# ================= UI =================
conversion_mode = st.selectbox(
    "📂 Pilih Jenis Konversi",
    [
        "PDF → PNG",
        "PDF → Word",
        "Word → PDF",
        "Excel → PDF"
    ]
)

dpi = st.selectbox("Resolusi DPI (PDF → PNG)", [150, 200, 300])

watermark = ""
if school_mode:
    school = st.text_input("Nama Sekolah")
    year = st.text_input("Tahun Ajaran", "2024/2025")
    watermark = f"{school} — Arsip Resmi — {year}"
else:
    watermark = st.text_input("Watermark (opsional)")

files = st.file_uploader(
    "Upload File",
    type=["pdf", "docx", "xlsx"],
    accept_multiple_files=True
)

# ================= PROCESS =================
if st.button("🚀 PROSES") and files:
    os.makedirs("output", exist_ok=True)
    bar = st.progress(0)

    for i, f in enumerate(files):
        path = save_temp(f)

        if conversion_mode == "PDF → PNG":
            out_dir = f"output/pdf_png_{i+1}"
            os.makedirs(out_dir, exist_ok=True)
            imgs = pdf_to_png(path, out_dir, dpi)
            if watermark:
                for img in imgs:
                    add_watermark(img, watermark)

        elif conversion_mode == "PDF → Word":
            out = path.replace(".pdf", ".docx")
            pdf_to_word(path, out)
            os.rename(out, f"output/{os.path.basename(out)}")

        elif conversion_mode == "Word → PDF":
            out = path.replace(".docx", ".pdf")
            word_to_pdf(path, out)
            os.rename(out, f"output/{os.path.basename(out)}")

        elif conversion_mode == "Excel → PDF":
            out = path.replace(".xlsx", ".pdf")
            excel_to_pdf(path, out)
            os.rename(out, f"output/{os.path.basename(out)}")

        bar.progress((i + 1) / len(files))

    zip_path = "HASIL_KONVERSI.zip"
    with zipfile.ZipFile(zip_path, "w") as z:
        for root, _, fs in os.walk("output"):
            for f in fs:
                full = os.path.join(root, f)
                z.write(full, arcname=full.replace("output/", ""))

    st.success("✅ Selesai")
    with open(zip_path, "rb") as z:
        st.download_button("📦 Download ZIP", z, file_name=zip_path)
