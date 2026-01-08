import streamlit as st
import os, zipfile
import fitz
from PIL import Image, ImageDraw
import pandas as pd
from pdf2docx import Converter
from docx2pdf import convert
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from rembg import remove

# ================= CONFIG =================
st.set_page_config(page_title="MassDoc Converter", layout="centered")
st.title("🧰 MassDoc Converter")
st.caption("Upload → Convert → Download")

st.divider()
school_mode = st.toggle("🏫 Aktifkan Mode Sekolah")
st.divider()

# ================= FUNCTIONS =================
def save_temp(file):
    path = f"temp_{file.name}"
    with open(path, "wb") as f:
        f.write(file.read())
    return path

def pdf_to_png(pdf_path, out_dir, dpi):
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    doc = fitz.open(pdf_path)
    imgs = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        out = f"{out_dir}/page_{i+1}.png"
        pix.save(out)
        imgs.append(out)
    return imgs

def png_to_pdf(images, out_pdf):
    imgs = [Image.open(i).convert("RGB") for i in images]
    imgs[0].save(out_pdf, save_all=True, append_images=imgs[1:])

def remove_bg(img_path, out_path):
    img = Image.open(img_path)
    result = remove(img)
    result.save(out_path)

def pdf_to_word(pdf, out):
    cv = Converter(pdf)
    cv.convert(out)
    cv.close()

def word_to_pdf(docx, out):
    convert(docx, out)

def excel_to_pdf(xlsx, out):
    df = pd.read_excel(xlsx)
    c = canvas.Canvas(out, pagesize=A4)
    w, h = A4
    y = h - 40
    for _, row in df.iterrows():
        x = 40
        for cell in row:
            c.drawString(x, y, str(cell))
            x += 100
        y -= 20
        if y < 40:
            c.showPage()
            y = h - 40
    c.save()

def watermark_img(img_path, text):
    img = Image.open(img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    draw.text((20, img.height - 40), text, fill=(150,150,150,120))
    img.save(img_path)

# ================= UI =================
mode = st.selectbox("📂 Pilih Mode Konversi", [
    "PDF → PNG",
    "PNG → PDF",
    "PNG → Remove Background",
    "PDF → Word",
    "Word → PDF",
    "Excel → PDF"
])

dpi = st.selectbox("Resolusi DPI", [150, 200, 300])

watermark = ""
if school_mode:
    school = st.text_input("Nama Sekolah")
    year = st.text_input("Tahun Ajaran", "2024/2025")
    watermark = f"{school} — Arsip Resmi — {year}"
else:
    watermark = st.text_input("Watermark (opsional)")

files = st.file_uploader(
    "Upload File",
    type=["pdf", "png", "docx", "xlsx"],
    accept_multiple_files=True
)

# ================= PROCESS =================
if st.button("🚀 PROSES") and files:
    os.makedirs("output", exist_ok=True)
    bar = st.progress(0)

    if mode == "PNG → PDF":
        img_paths = []
        for f in files:
            img_paths.append(save_temp(f))
        out_pdf = "output/PNG_to_PDF.pdf"
        png_to_pdf(img_paths, out_pdf)

    else:
        for i, f in enumerate(files):
            path = save_temp(f)

            if mode == "PDF → PNG":
                out_dir = f"output/pdf_png_{i}"
                os.makedirs(out_dir, exist_ok=True)
                imgs = pdf_to_png(path, out_dir, dpi)
                if watermark:
                    for img in imgs:
                        watermark_img(img, watermark)

            elif mode == "PNG → Remove Background":
                out = f"output/no_bg_{f.name}"
                remove_bg(path, out)

            elif mode == "PDF → Word":
                out = f"output/{f.name.replace('.pdf','.docx')}"
                pdf_to_word(path, out)

            elif mode == "Word → PDF":
                out = f"output/{f.name.replace('.docx','.pdf')}"
                word_to_pdf(path, out)

            elif mode == "Excel → PDF":
                out = f"output/{f.name.replace('.xlsx','.pdf')}"
                excel_to_pdf(path, out)

            bar.progress((i + 1) / len(files))

    zip_path = "HASIL_KONVERSI.zip"
    with zipfile.ZipFile(zip_path, "w") as z:
        for root, _, fs in os.walk("output"):
            for f in fs:
                full = os.path.join(root, f)
                z.write(full, arcname=full.replace("output/", ""))

    st.success("✅ Selesai")
    with open(zip_path, "rb") as f:
        st.download_button("📦 Download ZIP", f, file_name=zip_path)
