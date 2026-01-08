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

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Apiep Doc Converter",
    layout="centered"
)

# =====================================================
# AUTH CONFIG (LOGIN GURU)
# =====================================================
USERS = {
    "guru": "apiep123",
    "admin": "admin123"
}

if "login" not in st.session_state:
    st.session_state.login = False
    st.session_state.user = ""

def login_page():
    st.markdown("""
    <div class="glass">
    <h2>🔐 Login Guru</h2>
    <p>Gunakan akun resmi untuk mengakses aplikasi</p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")

        if submit:
            if u in USERS and USERS[u] == p:
                st.session_state.login = True
                st.session_state.user = u
                st.success("Login berhasil")
                st.rerun()
            else:
                st.error("Username / Password salah")

if not st.session_state.login:
    login_page()
    st.stop()

# =====================================================
# STYLE
# =====================================================
st.markdown("""
<style>
html, body {
    background: linear-gradient(120deg,#0f2027,#203a43,#2c5364);
}
.glass {
    background: rgba(255,255,255,0.12);
    backdrop-filter: blur(16px);
    border-radius: 24px;
    padding: 24px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.35);
    margin-bottom: 25px;
}
h1,h2,h3,label,p { color: white !important; }
.stButton>button {
    background: linear-gradient(90deg,#00c6ff,#0072ff);
    color:white; border-radius:14px;
    padding:0.7em 1.6em; font-weight:600;
}
[data-testid="stFileUploader"] {
    border:2px dashed rgba(255,255,255,0.4);
    border-radius:20px;
    padding:25px;
    background:rgba(255,255,255,0.05);
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# HEADER
# =====================================================
st.markdown(f"""
<div class="glass">
<h1>🧰 Apiep Doc Converter</h1>
<p>Login sebagai <b>{st.session_state.user}</b></p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# HELPERS
# =====================================================
def save_temp(file):
    path = f"temp_{file.name}"
    with open(path, "wb") as f:
        f.write(file.read())
    return path

def pdf_to_png(pdf, out_dir, dpi):
    zoom = dpi / 72
    doc = fitz.open(pdf)
    res = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        out = f"{out_dir}/page_{i+1}.png"
        pix.save(out)
        res.append(out)
    return res

def png_to_pdf(images, out_pdf):
    imgs = [Image.open(i).convert("RGB") for i in images]
    imgs[0].save(out_pdf, save_all=True, append_images=imgs[1:])

def remove_bg_img(img, out):
    remove(Image.open(img).convert("RGBA")).save(out, format="PNG")

def watermark_img(img_path, text):
    img = Image.open(img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    draw.text((20, img.height-40), text, fill=(180,180,180,120))
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

def pdf_to_word(pdf, out):
    c = Converter(pdf); c.convert(out); c.close()

def word_to_pdf(docx, out):
    convert(docx, out)

def excel_to_pdf(xlsx, out):
    df = pd.read_excel(xlsx)
    c = canvas.Canvas(out, pagesize=A4)
    y = A4[1] - 40
    for _, row in df.iterrows():
        x = 40
        for cell in row:
            c.drawString(x, y, str(cell))
            x += 100
        y -= 20
        if y < 40:
            c.showPage(); y = A4[1] - 40
    c.save()

# =====================================================
# UI
# =====================================================
st.markdown('<div class="glass">', unsafe_allow_html=True)

dpi = st.selectbox("Resolusi DPI", [150,200,300])
school_mode = st.toggle("🏫 Mode Sekolah")

if school_mode:
    school = st.text_input("Nama Sekolah")
    year = st.text_input("Tahun Ajaran", "2024/2025")
    watermark = f"{school} — {st.session_state.user.upper()} — {year}"
else:
    watermark = st.text_input("Watermark (opsional)")

files = st.file_uploader(
    "📤 Drag & Drop File",
    accept_multiple_files=True,
    type=["pdf","png","jpg","jpeg","docx","xlsx"]
)

mode = st.selectbox("📂 Mode Konversi", [
    "PDF → PNG",
    "PDF → Word",
    "PNG → PDF",
    "PNG → Remove Background",
    "Word → PDF",
    "Excel → PDF"
])

process = st.button("🚀 PROSES")
st.markdown('</div>', unsafe_allow_html=True)

# =====================================================
# PROCESS
# =====================================================
if process and files:
    os.makedirs("output", exist_ok=True)
    results = []
    bar = st.progress(0)

    for i, f in enumerate(files):
        path = save_temp(f)
        ext = os.path.splitext(f.name.lower())[1]

        if mode == "PDF → PNG" and ext == ".pdf":
            out_dir = "output/pdf_png"
            os.makedirs(out_dir, exist_ok=True)
            imgs = pdf_to_png(path, out_dir, dpi)
            if watermark:
                for img in imgs:
                    watermark_img(img, watermark)
            results.extend(imgs)

        elif mode == "PDF → Word" and ext == ".pdf":
            out = f"output/{f.name.replace('.pdf','.docx')}"
            pdf_to_word(path, out); results.append(out)

        elif mode == "PNG → PDF" and ext in [".png",".jpg",".jpeg"]:
            out = f"output/{f.name}.pdf"
            png_to_pdf([path], out); results.append(out)

        elif mode == "PNG → Remove Background" and ext in [".png",".jpg",".jpeg"]:
            out = f"output/no_bg_{f.name}.png"
            remove_bg_img(path, out); results.append(out)

        elif mode == "Word → PDF" and ext == ".docx":
            out = f"output/{f.name.replace('.docx','.pdf')}"
            word_to_pdf(path, out); results.append(out)

        elif mode == "Excel → PDF" and ext == ".xlsx":
            out = f"output/{f.name.replace('.xlsx','.pdf')}"
            excel_to_pdf(path, out); results.append(out)

        bar.progress((i+1)/len(files))

    if results:
        st.subheader("👀 Preview")
        if results[0].endswith(".png"):
            preview_images(results)
        else:
            preview_pdf(results[0])

        zip_path = "HASIL_KONVERSI.zip"
        with zipfile.ZipFile(zip_path,"w") as z:
            for r in results:
                z.write(r, arcname=os.path.basename(r))

        st.download_button("📦 Download ZIP", open(zip_path,"rb"), file_name=zip_path)
        st.success("🎉 Proses Selesai")
