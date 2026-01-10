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

# ================= SESSION STATE INIT =================
if "results" not in st.session_state:
    st.session_state.results = []

if "videos" not in st.session_state:
    st.session_state.videos = []

# ================= HEADER =================
st.markdown("## 🧰 Apiep Doc Converter\nKonversi Dokumen & Video Serba Praktis 🚀")

# ================= HELPERS =================
def save_temp(file):
    path = f"temp_{file.name}"
    with open(path, "wb") as f:
        f.write(file.read())
    return path

def safe_add_to_zip(zipf, filepath):
    if filepath and os.path.exists(filepath):
        zipf.write(filepath, arcname=os.path.basename(filepath))

# ================= CONVERTERS =================
def pdf_to_png(pdf, out_dir, dpi):
    zoom = dpi / 72
    doc = fitz.open(pdf)
    results = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        out = f"{out_dir}/page_{i+1}.png"
        pix.save(out)
        results.append(out)
    doc.close()
    return results

def preview_pdf_first_page(pdf_path, dpi=120):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
    img_path = pdf_path.replace(".pdf", "_preview.png")
    pix.save(img_path)
    doc.close()
    return img_path

def png_to_pdf(images, out_pdf):
    imgs = [Image.open(i).convert("RGB") for i in images]
    imgs[0].save(out_pdf, save_all=True, append_images=imgs[1:])

def pdf_to_word(pdf, out):
    c = Converter(pdf)
    c.convert(out)
    c.close()

def word_to_pdf(docx, out):
    out_dir = os.path.dirname(out)
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", docx, "--outdir", out_dir],
        check=True
    )
    base = os.path.splitext(os.path.basename(docx))[0]
    generated = os.path.join(out_dir, f"{base}.pdf")
    if generated != out and os.path.exists(generated):
        os.rename(generated, out)

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

    if res == "480p":
        clip = clip.resize(height=480)
    elif res == "720p":
        clip = clip.resize(height=720)
    elif res == "1080p":
        clip = clip.resize(height=1080)

    clip.write_videofile(
        out,
        codec="libx264",
        audio_codec="aac",
        preset="ultrafast",
        threads=2,
        logger=None
    )
    clip.close()

def jpg_to_png(img, out):
    Image.open(img).convert("RGBA").save(out, "PNG")

def png_to_jpg(img, out):
    im = Image.open(img)
    bg = Image.new("RGB", im.size, (255,255,255))
    if im.mode == "RGBA":
        bg.paste(im, mask=im.split()[3])
    else:
        bg.paste(im)
    bg.save(out, "JPEG", quality=85)

# ================= UI =================
mode = st.selectbox("📂 Mode Konversi", [
    "PDF → PNG","PDF → Word","PNG → PDF","Word → PDF",
    "Excel → PDF","JPG → PNG","PNG → JPG","MOV → MP4","AVI → MP4"
])

video_res = "Original"
if "MP4" in mode:
    video_res = st.selectbox("🎥 Resolusi Video", ["Original","480p","720p","1080p"])

dpi = st.selectbox("🖼️ Resolusi DPI", [150,200,300,600])

files = st.file_uploader(
    "📤 Upload File",
    accept_multiple_files=True,
    type=["pdf","png","jpg","jpeg","docx","xlsx","mov","avi"]
)

process = st.button("🚀 PROSES")

# ================= PROCESS =================
if process and files:
    os.makedirs("output", exist_ok=True)
    st.session_state.results.clear()
    st.session_state.videos.clear()

    for f in files:
        path = save_temp(f)
        ext = os.path.splitext(f.name.lower())[1]

        try:
            if mode=="PDF → PNG" and ext==".pdf":
                st.session_state.results += pdf_to_png(path,"output",dpi)

            elif mode=="PDF → Word" and ext==".pdf":
                out=f"output/{f.name.replace('.pdf','.docx')}"
                pdf_to_word(path,out)
                st.session_state.results.append(out)

            elif mode=="PNG → PDF" and ext in [".png",".jpg",".jpeg"]:
                out=f"output/{os.path.splitext(f.name)[0]}.pdf"
                png_to_pdf([path],out)
                st.session_state.results.append(out)

            elif mode=="Word → PDF" and ext==".docx":
                out=f"output/{f.name.replace('.docx','.pdf')}"
                word_to_pdf(path,out)
                st.session_state.results.append(out)

            elif mode=="Excel → PDF" and ext==".xlsx":
                out=f"output/{f.name.replace('.xlsx','.pdf')}"
                excel_to_pdf(path,out)
                st.session_state.results.append(out)

            elif mode=="JPG → PNG" and ext in [".jpg",".jpeg"]:
                out=f"output/{os.path.splitext(f.name)[0]}.png"
                jpg_to_png(path,out)
                st.session_state.results.append(out)

            elif mode=="PNG → JPG" and ext==".png":
                out=f"output/{os.path.splitext(f.name)[0]}.jpg"
                png_to_jpg(path,out)
                st.session_state.results.append(out)

            elif mode in ["MOV → MP4","AVI → MP4"] and ext in [".mov",".avi"]:
                out=f"output/{os.path.splitext(f.name)[0]}.mp4"
                video_to_mp4(path,out,video_res)
                st.session_state.videos.append(out)

        except Exception as e:
            st.error(f"❌ {f.name} gagal: {e}")

# ================= PREVIEW PDF =================
pdfs = [r for r in st.session_state.results if r.endswith(".pdf")]
if pdfs:
    st.subheader("📄 Preview PDF (Halaman 1)")
    for pdf in pdfs:
        img = preview_pdf_first_page(pdf)
        st.image(img, use_container_width=True)
        with open(pdf,"rb") as f:
            st.download_button("⬇️ Download PDF", f, file_name=os.path.basename(pdf))

# ================= DOWNLOAD ZIP =================
if st.session_state.results:
    zip_path="HASIL_KONVERSI.zip"
    with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as z:
        for r in st.session_state.results:
            safe_add_to_zip(z,r)

    with open(zip_path,"rb") as f:
        st.download_button("📦 Download ZIP", f, file_name=zip_path)

# ================= VIDEO =================
if st.session_state.videos:
    st.subheader("🎬 Preview & Download Video")
    for v in st.session_state.videos:
        st.video(v)
        with open(v,"rb") as f:
            st.download_button("⬇️ Download MP4", f, file_name=os.path.basename(v))
