import streamlit as st
import os
import fitz
from PIL import Image, ImageDraw
import zipfile

st.set_page_config(page_title="MassDoc Converter", layout="centered")

st.title("🧰 MassDoc Converter")
st.caption("Upload → Convert → Download")

st.divider()
school_mode = st.toggle("🏫 Aktifkan Mode Sekolah")
st.divider()

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

from pdf2docx import Converter

def pdf_to_word(pdf_path, output_docx):
    cv = Converter(pdf_path)
    cv.convert(output_docx)
    cv.close()

from docx2pdf import convert

def word_to_pdf(docx_path, output_pdf):
    convert(docx_path, output_pdf)

def add_watermark(img_path, text):
    img = Image.open(img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    draw.text((20, img.height - 40), text, fill=(150,150,150,120))
    img.save(img_path)

if school_mode:
    doc_type = st.selectbox("Jenis Dokumen Sekolah", ["Rapor", "Soal", "Nilai", "Surat"])
    preset = {
        "Rapor": (300, "RAPOR"),
        "Soal": (200, "SOAL"),
        "Nilai": (300, "NILAI"),
        "Surat": (300, "SURAT"),
    }
    dpi, prefix = preset[doc_type]
    school = st.text_input("Nama Sekolah")
    year = st.text_input("Tahun Ajaran", "2024/2025")
    watermark = f"{school} — Arsip Resmi — {year}"
else:
    dpi = st.selectbox("Resolusi DPI", [150, 200, 300])
    prefix = "FILE"
    watermark = st.text_input("Watermark (opsional)")

files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 PROSES") and files:
    os.makedirs("output", exist_ok=True)
    bar = st.progress(0)

    for i, f in enumerate(files):
        pdf_path = f"temp_{f.name}"
        with open(pdf_path, "wb") as out:
            out.write(f.read())

        out_dir = f"output/{prefix}_{i+1}"
        os.makedirs(out_dir, exist_ok=True)

        images = pdf_to_png(pdf_path, out_dir, dpi)
        if watermark:
            for img in images:
                add_watermark(img, watermark)

        bar.progress((i + 1) / len(files))

    zip_path = "HASIL_KONVERSI.zip"
    with zipfile.ZipFile(zip_path, "w") as z:
        for root, _, files in os.walk("output"):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, arcname=full.replace("output/", ""))

    st.success("✅ Selesai")
    with open(zip_path, "rb") as z:
        st.download_button("📦 Download ZIP", z, file_name=zip_path)

