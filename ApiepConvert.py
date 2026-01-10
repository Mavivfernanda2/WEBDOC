import streamlit as st
import os, zipfile
import fitz
import subprocess
from PIL import Image
import pandas as pd
from pdf2docx import Converter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from moviepy.editor import VideoFileClip

# ================= PAGE CONFIG =================
st.set_page_config(page_title="Apiep Doc Converter", layout="centered")

# ================= AUTH =================
USERS = {"guru": "apiep123", "admin": "admin123"}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = ""

def login_page():
    st.subheader("Selamat Datang")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u in USERS and USERS[u] == p:
            st.session_state.logged_in = True
            st.session_state.user = u
            st.rerun()
        else:
            st.error("❌ Username atau Password salah")

if not st.session_state.logged_in:
    login_page()
    st.stop()

# ================= HEADER =================
col1, col2 = st.columns([4,1])
with col1:
    st.markdown(f"## 🧰 Apiep Doc Converter\nLogin sebagai **{st.session_state.user.upper()}**")
with col2:
    if st.button("🚪 Logout"):
        st.session_state.logged_in = False
        st.session_state.user = ""
        st.rerun()

# ================= HELPERS =================
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

def pdf_to_word(pdf, out):
    c = Converter(pdf)
    c.convert(out)
    c.close()

def word_to_pdf(docx, out):
    output_dir = os.path.dirname(out)

    try:
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                docx,
                "--outdir",
                output_dir
            ],
            check=True
        )
    except FileNotFoundError:
        raise RuntimeError("LibreOffice tidak ditemukan. Gunakan Docker atau server dengan LibreOffice.")

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
            c.showPage()
            y = A4[1] - 40
    c.save()

def video_to_mp4(video_path, out_path, resolution):
    clip = VideoFileClip(video_path)

    if resolution == "480p":
        clip = clip.resize(height=480)
    elif resolution == "720p":
        clip = clip.resize(height=720)
    elif resolution == "1080p":
        clip = clip.resize(height=1080)

    clip.write_videofile(
        out_path,
        codec="libx264",
        audio_codec="aac",
        preset="ultrafast",
        threads=2
    )
    clip.close()

def jpg_to_png(img, out):
    Image.open(img).convert("RGBA").save(out, format="PNG")

def png_to_jpg(img, out, quality=85):
    im = Image.open(img)
    bg = Image.new("RGB", im.size, (255,255,255))
    bg.paste(im, mask=im.split()[3] if im.mode=="RGBA" else None)
    bg.save(out, "JPEG", quality=quality)

# ================= UI =================
mode = st.selectbox("📂 Mode Konversi", [
    "PDF → PNG",
    "PDF → Word",
    "PNG → PDF",
    "Word → PDF",
    "Excel → PDF",
    "JPG → PNG",
    "PNG → JPG",
    "MOV → MP4",
    "AVI → MP4"
])

video_res = "Original"
if mode in ["MOV → MP4", "AVI → MP4"]:
    video_res = st.selectbox("🎥 Resolusi Video", ["Original", "480p", "720p", "1080p"])

dpi = st.selectbox("Resolusi DPI", [150,200,300,600])

files = st.file_uploader(
    "📤 Upload File",
    accept_multiple_files=True,
    type=["pdf","png","jpg","jpeg","docx","xlsx","mov","avi"]
)

process = st.button("🚀 PROSES")

# ================= PROCESS =================
if process and files:
    os.makedirs("output", exist_ok=True)
    results, video_results = [], []
    bar = st.progress(0)

    for i, f in enumerate(files):
        path = save_temp(f)
        ext = os.path.splitext(f.name.lower())[1]

        if ext in [".mov",".avi"] and f.size > 300 * 1024 * 1024:
            st.error(f"❌ {f.name} terlalu besar (maks 300MB)")
            continue

        if mode == "PDF → PNG" and ext == ".pdf":
            results.extend(pdf_to_png(path, "output", dpi))

        elif mode == "PDF → Word" and ext == ".pdf":
            out = f"output/{f.name.replace('.pdf','.docx')}"
            pdf_to_word(path, out)
            results.append(out)

        elif mode == "PNG → PDF" and ext in [".png",".jpg",".jpeg"]:
            out = f"output/{f.name}.pdf"
            png_to_pdf([path], out)
            results.append(out)

        elif mode == "JPG → PNG" and ext in [".jpg",".jpeg"]:
            out = f"output/{os.path.splitext(f.name)[0]}.png"
            jpg_to_png(path, out)
            results.append(out)

        elif mode == "PNG → JPG" and ext == ".png":
            out = f"output/{os.path.splitext(f.name)[0]}.jpg"
            png_to_jpg(path, out)
            results.append(out)

        elif mode == "Word → PDF" and ext == ".docx":
            out = f"output/{f.name.replace('.docx','.pdf')}"
            word_to_pdf(path, out)
            results.append(out)

        elif mode == "Excel → PDF" and ext == ".xlsx":
            out = f"output/{f.name.replace('.xlsx','.pdf')}"
            excel_to_pdf(path, out)
            results.append(out)

        elif mode == "MOV → MP4" and ext == ".mov":
            out = f"output/{f.name.replace('.mov','.mp4')}"
            video_to_mp4(path, out, video_res)
            video_results.append(out)

        elif mode == "AVI → MP4" and ext == ".avi":
            out = f"output/{f.name.replace('.avi','.mp4')}"
            video_to_mp4(path, out, video_res)
            video_results.append(out)

        bar.progress((i+1)/len(files))

    if results:
        zip_path = "HASIL_KONVERSI.zip"
        with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as z:
            for r in results:
                z.write(r, arcname=os.path.basename(r))
        st.success("🎉 Proses Dokumen & Gambar Selesai")
        st.download_button("📦 Download ZIP", open(zip_path,"rb"), file_name=zip_path)

    if video_results:
        st.subheader("🎬 Download Video")
        for v in video_results:
            st.download_button(
                f"⬇️ {os.path.basename(v)}",
                open(v,"rb"),
                file_name=os.path.basename(v),
                mime="video/mp4"
            )
