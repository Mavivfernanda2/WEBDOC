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
st.set_page_config(
    page_title="Apiep Doc Converter",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ================= UI STYLE =================
st.markdown("""
<style>
body {
  background: radial-gradient(circle at top, #0f2027, #000);
}

.glass {
  background: rgba(255,255,255,0.08);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  border-radius: 22px;
  padding: 20px;
  border: 1px solid rgba(255,255,255,0.15);
  box-shadow: 0 8px 32px rgba(0,0,0,0.35);
  margin-bottom: 18px;
}

.header {
  background: linear-gradient(270deg,#00c6ff,#0072ff,#00ffd5);
  background-size: 600% 600%;
  animation: gradient 8s ease infinite;
  padding: 26px;
  border-radius: 28px;
  color: white;
  text-align: center;
  box-shadow: 0 10px 40px rgba(0,0,0,0.4);
}

.header h1 {
  margin-bottom: 6px;
}

@keyframes gradient {
  0% {background-position:0% 50%}
  50% {background-position:100% 50%}
  100% {background-position:0% 50%}
}
</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
st.markdown("""
<div class="header">
  <h1>🧰 Apiep Doc Converter</h1>
  <p>Convert • Preview • Download • Simple & Powerful</p>
</div>
""", unsafe_allow_html=True)

# ================= SESSION =================
if "results" not in st.session_state:
    st.session_state.results = []
if "videos" not in st.session_state:
    st.session_state.videos = []

# ================= HELPERS =================
def save_temp(file):
    path = f"temp_{file.name}"
    with open(path,"wb") as f:
        f.write(file.read())
    return path

def safe_add_to_zip(zipf, filepath):
    if os.path.exists(filepath):
        zipf.write(filepath, arcname=os.path.basename(filepath))

# ================= CONVERTERS =================
def preview_pdf_first_page(pdf_path, dpi=120):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
    img = pdf_path.replace(".pdf","_preview.png")
    pix.save(img)
    doc.close()
    return img

def pdf_to_png(pdf, out_dir, dpi):
    doc = fitz.open(pdf)
    res=[]
    zoom=dpi/72
    for i,p in enumerate(doc):
        pix=p.get_pixmap(matrix=fitz.Matrix(zoom,zoom))
        out=f"{out_dir}/page_{i+1}.png"
        pix.save(out)
        res.append(out)
    doc.close()
    return res

def pdf_to_word(pdf,out):
    c=Converter(pdf)
    c.convert(out)
    c.close()

def png_to_pdf(images,out):
    imgs=[Image.open(i).convert("RGB") for i in images]
    imgs[0].save(out,save_all=True,append_images=imgs[1:])

def word_to_pdf(docx,out):
    out_dir=os.path.dirname(out)
    subprocess.run(
        ["libreoffice","--headless","--convert-to","pdf",docx,"--outdir",out_dir],
        check=True
    )
    gen=os.path.join(out_dir,os.path.splitext(os.path.basename(docx))[0]+".pdf")
    if gen!=out:
        os.rename(gen,out)

def excel_to_pdf(xlsx,out):
    df=pd.read_excel(xlsx)
    c=canvas.Canvas(out,pagesize=A4)
    y=A4[1]-40
    for _,row in df.iterrows():
        x=40
        for cell in row:
            c.drawString(x,y,str(cell))
            x+=100
        y-=20
        if y<40:
            c.showPage()
            y=A4[1]-40
    c.save()

def video_to_mp4(video,out,res):
    clip=VideoFileClip(video)
    if res!="Original":
        clip=clip.resize(height=int(res.replace("p","")))
    clip.write_videofile(out,codec="libx264",audio_codec="aac",preset="ultrafast",logger=None)
    clip.close()

def jpg_to_png(img,out):
    Image.open(img).convert("RGBA").save(out,"PNG")

def png_to_jpg(img,out):
    im=Image.open(img)
    bg=Image.new("RGB",im.size,(255,255,255))
    bg.paste(im,mask=im.split()[3] if im.mode=="RGBA" else None)
    bg.save(out,"JPEG",quality=85)

# ================= UI =================
st.markdown('<div class="glass">', unsafe_allow_html=True)

mode = st.selectbox("⚙️ Mode Konversi", [...])
files = st.file_uploader("📤 Upload File", accept_multiple_files=True)
advanced = st.toggle("🎛 Advanced Mode")

if advanced:
    dpi=st.selectbox("🖼 DPI", [150,200,300,600])
    video_res=st.selectbox("🎥 Resolusi Video",["Original","480p","720p","1080p"])
else:
    dpi=150
    video_res="Original"

st.markdown('</div>', unsafe_allow_html=True)

# ================= PROCESS =================
if st.button("🚀 PROSES") and files:
    os.makedirs("output",exist_ok=True)
    st.session_state.results=[]
    st.session_state.videos=[]
    progress=st.progress(0)

    for i,f in enumerate(files):
        path=save_temp(f)
        ext=os.path.splitext(f.name.lower())[1]

        if mode=="PDF → PNG" and ext==".pdf":
            st.session_state.results+=pdf_to_png(path,"output",dpi)

        elif mode=="PDF → Word" and ext==".pdf":
            out=f"output/{f.name.replace('.pdf','.docx')}"
            pdf_to_word(path,out)
            st.session_state.results.append(out)

        elif mode=="PNG → PDF":
            out=f"output/{os.path.splitext(f.name)[0]}.pdf"
            png_to_pdf([path],out)
            st.session_state.results.append(out)

        elif mode=="Word → PDF" and ext==".docx":
            out=f"output/{f.name.replace('.docx','.pdf')}"
            word_to_pdf(path,out)
            st.session_state.results.append(out)

        elif mode=="Excel → PDF":
            out=f"output/{f.name.replace('.xlsx','.pdf')}"
            excel_to_pdf(path,out)
            st.session_state.results.append(out)

        elif mode=="JPG → PNG":
            out=f"output/{os.path.splitext(f.name)[0]}.png"
            jpg_to_png(path,out)
            st.session_state.results.append(out)

        elif mode=="PNG → JPG":
            out=f"output/{os.path.splitext(f.name)[0]}.jpg"
            png_to_jpg(path,out)
            st.session_state.results.append(out)

        elif mode in ["MOV → MP4","AVI → MP4"]:
            out=f"output/{os.path.splitext(f.name)[0]}.mp4"
            video_to_mp4(path,out,video_res)
            st.session_state.videos.append(out)

        progress.progress((i+1)/len(files))

    st.success("✅ Semua file selesai diproses")
    st.markdown("""
    <audio autoplay>
    <source src="https://assets.mixkit.co/sfx/preview/mixkit-click-error-1110.mp3">
    </audio>
    """,unsafe_allow_html=True)

# ================= PREVIEW =================
tab1,tab2=st.tabs(["👀 Preview","⬇️ Download"])

with tab1:
    for r in st.session_state.results:
        if r.endswith(".pdf"):
            img=preview_pdf_first_page(r)
            st.image(img,use_container_width=True)
    for v in st.session_state.videos:
        st.video(v)

with tab2:
    for r in st.session_state.results:
        st.download_button(f"⬇️ {os.path.basename(r)}",open(r,"rb"),file_name=os.path.basename(r))
    for v in st.session_state.videos:
        st.download_button(f"⬇️ {os.path.basename(v)}",open(v,"rb"),file_name=os.path.basename(v))

    if st.session_state.results:
        zip_path="HASIL_KONVERSI.zip"
        with zipfile.ZipFile(zip_path,"w") as z:
            for r in st.session_state.results:
                safe_add_to_zip(z,r)
        st.download_button("📦 Download ZIP",open(zip_path,"rb"),file_name=zip_path)
