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
st.set_page_config(page_title="Apiep Doc Converter", layout="centered")

# ================= STYLE =================
st.markdown("""
<style>
body {
    background: linear-gradient(120deg, #0f2027, #203a43, #2c5364);
}
img { animation: float 4s ease-in-out infinite; }
@keyframes float {
    0% { transform: translateY(0px); }
    50% { transform: translateY(-10px); }
    100% { transform: translateY(0px); }
}
.card {
    background: rgba(255,255,255,0.08);
    border-radius: 20px;
    padding: 20px;
    backdrop-filter: blur(10px);
}
h1, h2, h3, label { color: white !important; }
.stButton>button {
    background: linear-gradient(90deg, #00c6ff, #0072ff);
    color: white;
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
st.title("Apiep Doc Converter")
st.caption("Upload → Convert → Preview → Download")

st.markdown("""
<div class="card">
<h3>PDF • Word • Excel • Image</h3>
</div>
""", unsafe_allow_html=True)

st.divider()

# ================= MODE =================
school_mode = st.toggle("🏫 Mode Sekolah")
st.divider()

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
    result = remove(img)
    result.save(out_path, format="PNG")

def watermark_img(img_path, text):
    img = Image.open(img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    draw.text((20, img.height - 40), text, fill=(150,150,150,120))
    img.save(img_path)

def preview_images(imgs):
    cols = st.columns(3)
    for i, img in enumerate(imgs):
        with cols[i % 3]:
            st.image(img, use_container_width=True)

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
    results = []

    if mode == "PNG → PDF":
        imgs = [save_temp(f) for f in files]
        out_pdf = "output/PNG_to_PDF.pdf"
        png_to_pdf(imgs, out_pdf)
        st.success("✅ Berhasil")
        st.download_button("⬇️ Download PDF", open(out_pdf, "rb"), file_name="PNG_to_PDF.pdf")

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
                results.extend(imgs)

            elif mode == "PNG → Remove Background":
                out = f"output/no_bg_{os.path.splitext(f.name)[0]}.png"
                remove_bg_img(path, out)
                results.append(out)

            elif mode == "PDF → Word":
                out = f"output/{f.name.replace('.pdf','.docx')}"
                pdf_to_word(path, out)

            elif mode == "Word → PDF":
                out = f"output/{f.name.replace('.docx','.pdf')}"
                word_to_pdf(path, out)

            elif mode == "Excel → PDF":
                out = f"output/{f.name.replace('.xlsx','.pdf')}"
                excel_to_pdf(path, out)

        # ===== PREVIEW IMAGE =====
        if results:
            st.subheader("👀 Preview")
            preview_images(results)

            download_mode = st.radio(
                "📥 Opsi Download",
                ["Download PNG Satuan", "Download ZIP"],
                horizontal=True
            )

            if download_mode == "Download ZIP":
                zip_path = "HASIL_KONVERSI.zip"
                with zipfile.ZipFile(zip_path, "w") as z:
                    for r in results:
                        z.write(r, arcname=os.path.basename(r))
                st.download_button("📦 Download ZIP", open(zip_path, "rb"), file_name=zip_path)

            else:
                for r in results:
                    st.download_button(
                        f"⬇️ {os.path.basename(r)}",
                        open(r, "rb"),
                        file_name=os.path.basename(r),
                        mime="image/png"
                    )

        st.success("🎉 Proses Selesai")
