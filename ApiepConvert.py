import streamlit as st
import os, zipfile, subprocess
import fitz
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
    st.subheader("🔐 Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u in USERS and USERS[u] == p:
            st.session_state.logged_in = True
            st.session_state.user = u
            st.rerun()
        else:
            st.error("❌ Username / Password salah")

if not st.session_state.logged_in:
    login_page()
    st.stop()

# ================= HEADER =================
st.markdown(f"## 🧰 Apiep Doc Converter\nLogin sebagai **{st.session_state.user.upper()}**")
if st.button("🚪 Logout"):
    st.session_state.clear()
    st.rerun()

# ================= HELPERS =================
def save_temp(file):
    path = f"temp_{file.name}"
    with open(path, "wb") as f:
        f.write(file.read())
    return path

def safe_add_to_zip(zipf, filepath):
    if filepath and os.path.exists(filepath):
        zipf.write(filepath, arcname=os.path.basename(filepath))

def preview_pdf_first_page(pdf_path, dpi=120):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
    img_path = pdf_path.replace(".pdf", "_preview.png")
    pix.save(img_path)
    doc.close()
    return img_path

# ================= CONVERTERS =================
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
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", docx, "--outdir", "output"],
        check=True
    )
    gen = f"output/{os.path.splitext(os.path.basename(docx))[0]}.pdf"
    os.rename(gen, out)

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

def video_to_mp4(video, out, res):
    clip = VideoFileClip(video)
    if res != "Original":
        clip = clip.resize(height=int(res.replace("p","")))
    clip.write_videofile(out, codec="libx264", audio_codec="aac", preset="ultrafast", logger=None)
    clip.close()

def jpg_to_png(img, out):
    Image.open(img).convert("RGBA").save(out)

def png_to_jpg(img, out):
    im = Image.open(img).convert("RGB")
    im.save(out, "JPEG", quality=85)

# ================= UI =================
mode = st.selectbox("📂 Mode Konversi", [
    "PDF → PNG","PDF → Word","PNG → PDF","Word → PDF","Excel → PDF",
    "JPG → PNG","PNG → JPG","MOV → MP4","AVI → MP4"
])

video_res = "Original"
if "MP4" in mode:
    video_res = st.selectbox("🎥 Resolusi Video", ["Original","480p","720p","1080p"])

dpi = st.selectbox("🖼️ DPI (PDF → PNG)", [150,200,300])

files = st.file_uploader(
    "📤 Upload File",
    accept_multiple_files=True,
    type=["pdf","png","jpg","jpeg","docx","xlsx","mov","avi"]
)

process = st.button("🚀 PROSES")

# ================= PROCESS =================
if process and files:
    os.makedirs("output", exist_ok=True)
    results, videos = [], []
    bar = st.progress(0)

    for i, f in enumerate(files):
        path = save_temp(f)
        ext = os.path.splitext(f.name.lower())[1]

        try:
            if mode == "PDF → PNG" and ext == ".pdf":
                results += pdf_to_png(path, "output", dpi)

            elif mode == "PDF → Word" and ext == ".pdf":
                out = f"output/{f.name.replace('.pdf','.docx')}"
                pdf_to_word(path, out)
                results.append(out)

            elif mode == "PNG → PDF":
                out = f"output/{os.path.splitext(f.name)[0]}.pdf"
                png_to_pdf([path], out)
                results.append(out)

            elif mode == "Word → PDF" and ext == ".docx":
                out = f"output/{f.name.replace('.docx','.pdf')}"
                word_to_pdf(path, out)
                results.append(out)

            elif mode == "Excel → PDF":
                out = f"output/{f.name.replace('.xlsx','.pdf')}"
                excel_to_pdf(path, out)
                results.append(out)

            elif mode == "JPG → PNG":
                out = f"output/{os.path.splitext(f.name)[0]}.png"
                jpg_to_png(path, out)
                results.append(out)

            elif mode == "PNG → JPG":
                out = f"output/{os.path.splitext(f.name)[0]}.jpg"
                png_to_jpg(path, out)
                results.append(out)

            elif "MP4" in mode:
                out = f"output/{os.path.splitext(f.name)[0]}.mp4"
                video_to_mp4(path, out, video_res)
                videos.append(out)

        except Exception as e:
            st.error(f"❌ {f.name} gagal: {e}")

        bar.progress((i+1)/len(files))

# ================= PREVIEW PDF =================
if results:
    pdfs = [r for r in results if r.endswith(".pdf")]
    if pdfs:
        st.subheader("📄 Preview PDF (Halaman 1)")
        for pdf in pdfs:
            img = preview_pdf_first_page(pdf)
            st.image(img, use_container_width=True)
            st.download_button("⬇️ Download PDF", open(pdf,"rb"), file_name=os.path.basename(pdf))

# ================= DOWNLOAD FILE =================
if results:
    zip_path = "HASIL_KONVERSI.zip"
    with zipfile.ZipFile(zip_path,"w") as z:
        for r in results:
            safe_add_to_zip(z, r)

    st.download_button("📦 Download ZIP", open(zip_path,"rb"), file_name=zip_path)

# ================= VIDEO =================
if videos:
    st.subheader("🎬 Preview & Download Video")
    for v in videos:
        st.video(v)
        st.download_button("⬇️ Download MP4", open(v,"rb"), file_name=os.path.basename(v))
