{
  "nbformat": 4,
  "nbformat_minor": 0,
  "metadata": {
    "colab": {
      "provenance": [],
      "authorship_tag": "ABX9TyMl+AWA6WZJx6XAELhyTjqj",
      "include_colab_link": True
    },
    "kernelspec": {
      "name": "python3",
      "display_name": "Python 3"
    },
    "language_info": {
      "name": "python"
    }
  },
  "cells": [
    {
      "cell_type": "markdown",
      "metadata": {
        "id": "view-in-github",
        "colab_type": "text"
      },
      "source": [
        "<a href=\"https://colab.research.google.com/github/Mavivfernanda2/WEBDOC/blob/main/MassDoc_Final.py\" target=\"_parent\"><img src=\"https://colab.research.google.com/assets/colab-badge.svg\" alt=\"Open In Colab\"/></a>"
      ]
    },
    {
      "cell_type": "code",
      "execution_count": null,
      "metadata": {
        "id": "sYEtQrwpg4HO"
      },
      "outputs": [],
      "source": [
        "!pip install streamlit pymupdf pillow\n"
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "import os\n",
        "\n",
        "os.makedirs(\"utils\", exist_ok=True)\n",
        "os.makedirs(\"output\", exist_ok=True)\n"
      ],
      "metadata": {
        "id": "B4Lj6W6sg_-A"
      },
      "execution_count": 3,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "%%writefile utils/converter.py\n",
        "import fitz\n",
        "import os\n",
        "\n",
        "def pdf_to_png(pdf_path, output_dir, dpi=300):\n",
        "    doc = fitz.open(pdf_path)\n",
        "    images = []\n",
        "\n",
        "    zoom = dpi / 72\n",
        "    mat = fitz.Matrix(zoom, zoom)\n",
        "\n",
        "    for i, page in enumerate(doc):\n",
        "        pix = page.get_pixmap(matrix=mat)\n",
        "        out = os.path.join(output_dir, f\"page_{i+1}.png\")\n",
        "        pix.save(out)\n",
        "        images.append(out)\n",
        "\n",
        "    return images\n"
      ],
      "metadata": {
        "id": "zHI70JXVhCSA"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "%%writefile utils/school.py\n",
        "def school_preset(doc_type):\n",
        "    presets = {\n",
        "        \"Rapor\": {\"dpi\":300, \"wm\":True, \"prefix\":\"RAPOR\"},\n",
        "        \"Soal Ujian\": {\"dpi\":200, \"wm\":False, \"prefix\":\"SOAL\"},\n",
        "        \"Nilai\": {\"dpi\":300, \"wm\":True, \"prefix\":\"NILAI\"},\n",
        "        \"Surat Resmi\": {\"dpi\":300, \"wm\":True, \"prefix\":\"SURAT\"}\n",
        "    }\n",
        "    return presets[doc_type]\n"
      ],
      "metadata": {
        "id": "Xi4AKHclhFa9"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "%%writefile utils/watermark.py\n",
        "from PIL import Image, ImageDraw\n",
        "\n",
        "def add_watermark(img_path, text):\n",
        "    img = Image.open(img_path).convert(\"RGBA\")\n",
        "    draw = ImageDraw.Draw(img)\n",
        "\n",
        "    draw.text(\n",
        "        (20, img.height - 40),\n",
        "        text,\n",
        "        fill=(150,150,150,120)\n",
        "    )\n",
        "\n",
        "    img.save(img_path)\n"
      ],
      "metadata": {
        "id": "j8aqhlrkhIjY"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "%%writefile utils/zipper.py\n",
        "import zipfile\n",
        "import os\n",
        "\n",
        "def make_zip(folder, zip_path):\n",
        "    with zipfile.ZipFile(zip_path, \"w\") as z:\n",
        "        for root, _, files in os.walk(folder):\n",
        "            for f in files:\n",
        "                full = os.path.join(root, f)\n",
        "                z.write(full, arcname=full.replace(folder+\"/\",\"\"))\n"
      ],
      "metadata": {
        "id": "4NFkOyCLhK1o"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "%%writefile app.py\n",
        "import streamlit as st\n",
        "import os\n",
        "from utils.converter import pdf_to_png\n",
        "from utils.school import school_preset\n",
        "from utils.watermark import add_watermark\n",
        "from utils.zipper import make_zip\n",
        "\n",
        "st.set_page_config(page_title=\"MassDoc Converter\", layout=\"centered\")\n",
        "\n",
        "st.title(\"🧰 MassDoc Converter\")\n",
        "st.caption(\"Upload → Convert → Download\")\n",
        "\n",
        "st.divider()\n",
        "school_mode = st.toggle(\"🏫 Aktifkan Mode Sekolah\")\n",
        "st.divider()\n",
        "\n",
        "if school_mode:\n",
        "    doc_type = st.selectbox(\n",
        "        \"Jenis Dokumen Sekolah\",\n",
        "        [\"Rapor\", \"Soal Ujian\", \"Nilai\", \"Surat Resmi\"]\n",
        "    )\n",
        "    preset = school_preset(doc_type)\n",
        "    dpi = preset[\"dpi\"]\n",
        "    prefix = preset[\"prefix\"]\n",
        "\n",
        "    school_name = st.text_input(\"Nama Sekolah\")\n",
        "    tahun = st.text_input(\"Tahun Ajaran\", \"2024/2025\")\n",
        "    watermark_text = f\"{school_name} — Arsip Resmi — {tahun}\"\n",
        "\n",
        "    st.info(f\"DPI otomatis: {dpi}\")\n",
        "else:\n",
        "    dpi = st.selectbox(\"Resolusi (DPI)\", [150,200,300])\n",
        "    watermark_text = st.text_input(\"Watermark (opsional)\")\n",
        "    prefix = \"FILE\"\n",
        "\n",
        "files = st.file_uploader(\n",
        "    \"Upload PDF\",\n",
        "    type=[\"pdf\"],\n",
        "    accept_multiple_files=True\n",
        ")\n",
        "\n",
        "if st.button(\"🚀 PROSES\") and files:\n",
        "    os.makedirs(\"output\", exist_ok=True)\n",
        "\n",
        "    bar = st.progress(0)\n",
        "    status = st.empty()\n",
        "\n",
        "    for i, f in enumerate(files):\n",
        "        status.write(f\"📄 {f.name}\")\n",
        "\n",
        "        pdf_path = f\"temp_{f.name}\"\n",
        "        with open(pdf_path,\"wb\") as out:\n",
        "            out.write(f.read())\n",
        "\n",
        "        out_dir = f\"output/{prefix}_{i+1}\"\n",
        "        os.makedirs(out_dir, exist_ok=True)\n",
        "\n",
        "        images = pdf_to_png(pdf_path, out_dir, dpi)\n",
        "\n",
        "        if watermark_text:\n",
        "            for img in images:\n",
        "                add_watermark(img, watermark_text)\n",
        "\n",
        "        bar.progress((i+1)/len(files))\n",
        "\n",
        "    zip_path = \"HASIL_KONVERSI.zip\"\n",
        "    make_zip(\"output\", zip_path)\n",
        "\n",
        "    st.success(\"✅ Selesai\")\n",
        "\n",
        "    with open(zip_path,\"rb\") as z:\n",
        "        st.download_button(\n",
        "            \"📦 Download ZIP\",\n",
        "            z,\n",
        "            file_name=\"HASIL_KONVERSI.zip\"\n",
        "        )\n"
      ],
      "metadata": {
        "id": "IR-8yBAxhMgV"
      },
      "execution_count": null,
      "outputs": []
    }
  ]
}
