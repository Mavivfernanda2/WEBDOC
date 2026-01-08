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
h1,h2,h3,label { color:white !important; }
.stButton>button {
    background: linear-gradient(90deg,#00c6ff,#0072ff);
    color:white;
    border-radius:12px;
}
.card {
    background: rgba(255,255,255,0.08);
    border-radius:20px;
    padding:20px;
    backdrop-filter: blur(10px);
}
</style>
""", unsafe_allow_html=True)

# ================= CONFIG =================
st.set_page_config("MassDoc Converter", layout="centered")
st.title("🧰 MassDoc Converter")
st.caption("Upload → Convert → Preview → Download")

st.markdown("""
<div class="card">
<h2>📄 PDF • Word • Excel • Image</h2>
<img src="https://raw.githubusercontent.com/edent/SuperTinyIcons/master/images/svg/pdf.svg" width="60">
<img src="https://raw.githubusercontent.com/edent/SuperTinyIcons/master/images/svg/microsoftword.svg" width="60">
<img src="https://raw.githubusercontent.com/edent/SuperTinyIcons/master/images/svg/microsoftexcel.svg" width="60">
<img src="https://raw.githubusercontent.com/edent/SuperTinyIcons/master/images/svg/image.svg" width="60">
</div>
""", unsafe_allow_html=True)

st.divider()
school_mode = st.toggle("🏫 Mode Sekolah")
st.divider()

# ================= FUNCTIONS =================
def save_temp(f):
    p = f"temp_{f.name}"
    with open(p, "wb") as o:
        o.write(f.read())
    return p

def pdf_to_png(pdf, out_dir, dpi):
    os.makedirs(out_dir, exist_ok=True)
    zoom = dpi / 72
    doc = fitz.open(pdf)
    imgs = []
    for i,p in enumerate(doc):
        pix = p.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        out = f"{out_dir}/page_{i+1}.png"
        pix.save(out)
        imgs.append(out)
    return imgs

def png_to_pdf(imgs, out):
    images = [Image.open(i).convert("RGB") for i in imgs]
    images[0].save(out, save_all=True, append_images=images[1:])

def remove_bg_img(src, out):
    img = Image.open(src)
    remove(img).save(out)

def pdf_to_word(pdf, out):
    cv = Converter(pdf); cv.convert(out); cv.close()

def word_to_pdf(docx, out):
    convert(docx, out)

def excel_to_pdf(xlsx, out):
    df = pd.read_excel(xlsx)
    c = canvas.Canvas(out, pagesize=A4)
    y = A4[1] - 40
    for _,r in df.iterrows():
        x = 40
        for cell in r:
            c.drawString(x, y, str(cell)); x+=100
        y -= 20
        if y < 40:
            c.showPage(); y = A4[1] - 40
    c.save()

def preview_images(imgs):
    cols = st.columns(3)
    for i,img in enumerate(imgs):
        with cols[i%3]:
            st.image(img, use_container_width=True)

# ================= UI =================
mode = st.selectbox("Pilih Mode", [
    "PDF → PNG", "PNG → PDF", "PNG → Remove Background",
    "PDF → Word", "Word → PDF", "Excel → PDF"
])

dpi = st.selectbox("Resolusi DPI", [150,200,300])
files = st.file_uploader("Upload File", accept_multiple_files=True)

# ================= PROCESS =================
if st.button("🚀 PROSES") and files:
    os.makedirs("output", exist_ok=True)
    results = []

    if mode == "PNG → PDF":
        imgs = [save_temp(f) for f in files]
        out = "output/PNG_to_PDF.pdf"
        png_to_pdf(imgs, out)
        results.append(out)

    else:
        for f in files:
            p = save_temp(f)

            if mode == "PDF → PNG":
                results += pdf_to_png(p, "output/png", dpi)

            elif mode == "PNG → Remove Background":
                out = f"output/no_bg_{f.name}"
                remove_bg_img(p, out)
                results.append(out)

            elif mode == "PDF → Word":
                out = f"output/{f.name}.docx"
                pdf_to_word(p, out)
                results.append(out)

            elif mode == "Word → PDF":
                out = f"output/{f.name}.pdf"
                word_to_pdf(p, out)
                results.append(out)

            elif mode == "Excel → PDF":
                out = f"output/{f.name}.pdf"
                excel_to_pdf(p, out)
                results.append(out)

    # ================= PREVIEW =================
    st.subheader("👀 Preview")
    if any(r.endswith(".png") for r in results):
        preview_images(results)

    # ================= DOWNLOAD =================
    if len(results) == 1 and results[0].endswith(".png"):
        with open(results[0],"rb") as f:
            st.download_button("⬇️ Download PNG", f, file_name=os.path.basename(results[0]))
    else:
        zipf = "HASIL_KONVERSI.zip"
        with zipfile.ZipFile(zipf,"w") as z:
            for r in results:
                z.write(r, arcname=os.path.basename(r))
        with open(zipf,"rb") as f:
            st.download_button("📦 Download ZIP", f, file_name=zipf)
