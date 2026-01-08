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

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="Apiep Doc Converter",
    layout="centered"
)

# ================= STYLE =================
st.markdown("""
<style>
html, body {
    background: linear-gradient(120deg, #0f2027, #203a43, #2c5364);
}
.glass {
    background: rgba(255,255,255,0.12);
    backdrop-filter: blur(16px);
    border-radius: 24px;
    padding: 24px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.35);
    margin-bottom: 25px;
}
h1,h2,h3,label,p {
    color: white !important;
}
.stButton > button {
    background: linear-gradient(90deg,#00c6ff,#0072ff);
    color: white;
    border-radius: 14px;
    padding: 0.7em 1.6em;
    font-weight: 600;
}
[data-testid="stFileUploader"] {
    border: 2px dashed rgba(255,255,255,0.4);
    border-radius: 20px;
    padding: 25px;
    background: rgba(255,255,255,0.05);
}
</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
st.markdown("""
<div class="glass">
<h1>🧰 Apiep Doc Converter</h1>
<p>Upload → Convert → Preview → Download</p>
</div>
""", unsafe_allow_html=True)

# ================= HELPERS =================
def save_temp(file):
    path = f"temp_{file.name}"
    with open(path, "wb") as f:
        f.write(file.read())
    return path

def pdf_to_png(pdf, out_dir, dpi):
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    doc = fitz.open(pdf)
    results = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        out = f"{out_dir}/page_{i+1}.png"
        pix.save(out)
        results.append(out)
    return results

def png_to_pdf(images, out_pdf):
    imgs = [Image.open(i).convert("RGB") for i in images]
    imgs[0].save(out_pdf, save_all=True, append_images=imgs[1:])

def remove_bg_img(img_path, out_path):
    img = Image.open(img_path).convert("RGBA")
    remove(img).save(out_path, format="PNG")

def watermark_img(img_path, text):
    img = Image.open(img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    draw.text((20, img.height - 40), text, fill=(180,180,180,120))
    img.save(img_path)

def preview_images(imgs):
    cols = st.columns(3)
    for i, img in enumerate(imgs):
        with cols[i % 3]:
            st.image(img, use_container_width=True)

def preview_pdf(path):
    with open(path, "rb") as f:
        st.download_button(
            "👀 Preview / Download PDF",
            f,
            file_name=os.path.basename(path),
            mime="application/pdf"
        )

# ================= MODE SELECTION (WRAP) =================
st.markdown('<div class="glass">', unsafe_allow_html=True)

mode = st.selectbox(
    "📂 Pilih Mode Konversi",
    [
        "PDF → PNG",
        "PNG → PDF",
        "PNG → Remove Background",
        "PDF → Word",
        "Word → PDF",
        "Excel → PDF"
    ]
)

dpi = st.selectbox("Resolusi DPI", [150, 200, 300])
school_mode = st.toggle("🏫 Mode Sekolah")

if school_mode:
    school = st.text_input("Nama Sekolah")
    year = st.text_input("Tahun Ajaran", "2024/2025")
    watermark = f"{school} — Arsip Resmi — {year}"
else:
    watermark = st.text_input("Watermark (opsional)")

st.markdown('</div>', unsafe_allow_html=True)

# ================= UPLOAD AREA (WRAP) =================
st.markdown("""
<div class="glass">
<h3>📤 Upload File</h3>
<p>Drag & drop file kamu di sini</p>
</div>
""", unsafe_allow_html=True)

files = st.file_uploader(
    "",
    type=["pdf","png","docx","xlsx"],
    accept_multiple_files=True
)

process = st.button("🚀 PROSES")

# ================= PROCESS =================
if process and files:
    os.makedirs("output", exist_ok=True)
    results = []

    if mode == "PNG → PDF":
        imgs = [save_temp(f) for f in files]
        out_pdf = "output/PNG_to_PDF.pdf"
        png_to_pdf(imgs, out_pdf)
        preview_pdf(out_pdf)
        results.append(out_pdf)

    else:
        for f in files:
            path = save_temp(f)

            if mode == "PDF → PNG":
                out_dir = f"output/pdf_png"
                os.makedirs(out_dir, exist_ok=True)
                imgs = pdf_to_png(path, out_dir, dpi)
                if watermark:
                    for img in imgs:
                        watermark_img(img, watermark)
                results.extend(imgs)

            elif mode == "PNG → Remove Background":
                out = f"output/no_bg_{f.name}"
                remove_bg_img(path, out)
                results.append(out)

    # ===== PREVIEW =====
    if results:
        st.subheader("👀 Preview")
        if results[0].endswith(".png"):
            preview_images(results)
        else:
            preview_pdf(results[0])

        zip_path = "HASIL_KONVERSI.zip"
        with zipfile.ZipFile(zip_path, "w") as z:
            for r in results:
                z.write(r, arcname=os.path.basename(r))

        st.download_button(
            "📦 Download ZIP",
            open(zip_path, "rb"),
            file_name=zip_path
        )

        st.success("🎉 Proses Selesai")
