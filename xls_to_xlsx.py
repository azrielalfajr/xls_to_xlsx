import streamlit as st
import pandas as pd
import io
import zipfile
import os
import numpy as np
from io import StringIO
from tempfile import TemporaryDirectory
from email import message_from_bytes
from email.policy import default as email_default_policy

st.set_page_config(page_title="XLS to XLSX Converter", layout="centered")
st.title("📄 Konversi XLS ke XLSX dan Download ZIP")
st.write(
    "Upload beberapa file .xls dan akan dikonversi ke .xlsx. "
    "Baris pertama sampai ke-8 akan otomatis dihapus. "
    "Semua file hasil akan diunduh dalam bentuk .zip."
)

uploaded_files = st.file_uploader("Upload file .xls", type=["xls"], accept_multiple_files=True)

# -------- Helpers --------
def is_probably_html(content: bytes) -> bool:
    head = content[:4096].lower()
    return (b"<html" in head) or (b"<!doctype html" in head) or (b"<table" in head)

def is_probably_mhtml(content: bytes) -> bool:
    head = content[:4096]
    # Ciri umum file Internet Explorer "Save as Web Archive" atau export dari aplikasi
    return head.startswith(b"MIME-Version:") or b"multipart/related" in head[:2048].lower()

def extract_html_from_mhtml(content: bytes) -> str:
    """Ambil bagian text/html terbesar dari berkas MHTML."""
    msg = message_from_bytes(content, policy=email_default_policy)
    html_parts = []

    def walk(m):
        if m.is_multipart():
            for part in m.iter_parts():
                walk(part)
        else:
            if m.get_content_type() == "text/html":
                try:
                    html_parts.append(m.get_content())
                except Exception:
                    pass

    walk(msg)
    if not html_parts:
        raise ValueError("Tidak menemukan bagian text/html di MHTML.")
    # Ambil yang paling besar (biasanya tabel utama)
    return max(html_parts, key=len)

def read_html_table(html_text: str) -> pd.DataFrame:
    dfs = pd.read_html(StringIO(html_text))
    if not dfs:
        raise ValueError("Tidak ditemukan <table> pada HTML.")
    return dfs[0]

def clean_after_skip(df: pd.DataFrame) -> pd.DataFrame:
    # Ganti string kosong jadi NaN agar mudah dibuang
    df = df.replace("", np.nan)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df = df.reset_index(drop=True)
    return df

def convert_to_dataframe(uploaded_file):
    """
    Strategi:
      - Jika MHTML: ekstrak HTML → read_html → skip 8 baris manual
      - else jika HTML polos: read_html → skip 8 baris manual
      - else: coba XLS biner (xlrd==1.2.0) dengan skiprows=8
      - fallback silang kalau deteksi awal keliru
    """
    content = uploaded_file.read()
    uploaded_file.seek(0)

    # 1) MHTML?
    if is_probably_mhtml(content):
        try:
            html_text = extract_html_from_mhtml(content)
            df = read_html_table(html_text)
            df = df.iloc[8:] if len(df) > 8 else df.iloc[0:0]
            df = clean_after_skip(df)
            return df, None
        except Exception as e:
            # fallback ke Excel atau HTML polos
            mhtml_err = e

            # Coba Excel biner
            try:
                df = pd.read_excel(io.BytesIO(content), engine="xlrd", skiprows=8)
                df = clean_after_skip(df)
                return df, None
            except Exception as e_xls:
                # Coba HTML polos
                try:
                    text = content.decode("utf-8", errors="ignore")
                    df = read_html_table(text)
                    df = df.iloc[8:] if len(df) > 8 else df.iloc[0:0]
                    df = clean_after_skip(df)
                    return df, None
                except Exception as e_html:
                    return None, f"Gagal baca MHTML/Excel/HTML. MHTML: {mhtml_err}; Excel: {e_xls}; HTML: {e_html}"

    # 2) HTML polos?
    if is_probably_html(content):
        try:
            text = content.decode("utf-8", errors="ignore")
            df = read_html_table(text)
            df = df.iloc[8:] if len(df) > 8 else df.iloc[0:0]
            df = clean_after_skip(df)
            return df, None
        except Exception as e_html:
            # fallback Excel
            try:
                df = pd.read_excel(io.BytesIO(content), engine="xlrd", skiprows=8)
                df = clean_after_skip(df)
                return df, None
            except Exception as e_xls:
                return None, f"Gagal baca HTML maupun Excel. HTML: {e_html}; Excel: {e_xls}"

    # 3) Asumsikan XLS biner
    try:
        df = pd.read_excel(io.BytesIO(content), engine="xlrd", skiprows=8)
        df = clean_after_skip(df)
        return df, None
    except Exception as e_xls:
        # fallback terakhir: coba treat sebagai HTML (kadang signature tidak tampak di awal)
        try:
            text = content.decode("utf-8", errors="ignore")
            df = read_html_table(text)
            df = df.iloc[8:] if len(df) > 8 else df.iloc[0:0]
            df = clean_after_skip(df)
            return df, None
        except Exception as e_html:
            return None, f"Tidak bisa baca file sebagai XLS maupun HTML. XLS: {e_xls}; HTML: {e_html}"

# -------- Main --------
if uploaded_files:
    if st.button("🔄 Konversi dan Download ZIP"):
        with TemporaryDirectory() as tmpdir:
            zip_buffer = io.BytesIO()
            success_count = 0

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for uploaded_file in uploaded_files:
                    with st.spinner(f"Memproses {uploaded_file.name}..."):
                        df, error = convert_to_dataframe(uploaded_file)

                    if error:
                        st.error(f"❌ {uploaded_file.name}: {error}")
                        continue

                    new_filename = os.path.splitext(uploaded_file.name)[0] + ".xlsx"
                    output_path = os.path.join(tmpdir, new_filename)

                    try:
                        df.to_excel(output_path, index=False, engine="openpyxl")
                        zipf.write(output_path, arcname=new_filename)
                        success_count += 1
                        st.success(f"✅ Berhasil: {new_filename}")
                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan {new_filename}: {e}")

            if success_count > 0:
                zip_buffer.seek(0)
                st.download_button(
                    label=f"📥 Download hasil ({success_count} file) (.zip)",
                    data=zip_buffer,
                    file_name="converted_xlsx_files.zip",
                    mime="application/zip"
                )
            else:
                st.warning("Tidak ada file yang berhasil dikonversi.")
else:
    st.info("Silakan upload satu atau lebih file .xls terlebih dahulu.")
