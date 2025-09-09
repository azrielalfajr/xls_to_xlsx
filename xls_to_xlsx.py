import streamlit as st
import pandas as pd
import io
import zipfile
import os
from io import StringIO
from tempfile import TemporaryDirectory

# -----------------------------
# App setup
# -----------------------------
st.set_page_config(page_title="XLS to XLSX Converter", layout="centered")
st.title("📄 Konversi XLS ke XLSX dan Download ZIP")
st.write(
    "Upload beberapa file .xls dan akan dikonversi ke .xlsx. "
    "Baris pertama sampai ke-8 akan otomatis dihapus. "
    "Semua file hasil akan diunduh dalam bentuk .zip."
)

uploaded_files = st.file_uploader("Upload file .xls", type=["xls"], accept_multiple_files=True)

# -----------------------------
# Helpers
# -----------------------------
def is_probably_html(content_bytes: bytes) -> bool:
    """
    Banyak file .xls dari sistem legacy sebenarnya HTML yang bisa dibuka Excel.
    Deteksi tanda-tanda HTML agar dibaca via pandas.read_html.
    """
    head = content_bytes[:4096].lower()
    return (
        b"<html" in head
        or b"<!doctype html" in head
        or b"<table" in head
    )

def try_read_html_to_df(content_bytes: bytes):
    """
    Baca konten HTML dan kembalikan DataFrame pertama.
    """
    text = content_bytes.decode("utf-8", errors="ignore")
    dfs = pd.read_html(StringIO(text))  # butuh lxml/bs4/html5lib
    if not dfs:
        return None, "Tidak ditemukan tabel dalam file HTML."
    return dfs[0], None

def try_read_xls_to_df(content_bytes: bytes, skiprows=8):
    """
    Baca .xls biner menggunakan xlrd (perlu xlrd==1.2.0).
    Return DF pertama (atau DF sheet pertama).
    """
    # Jika multi-sheet dan kamu ingin sheet tertentu, bisa tambahkan sheet_name=
    df = pd.read_excel(io.BytesIO(content_bytes), engine="xlrd", skiprows=skiprows)
    return df

def clean_df_after_skip(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pembersihan ringan pasca skip baris:
    - Reset index
    - Drop kolom kosong total
    """
    df = df.reset_index(drop=True)
    df = df.dropna(axis=1, how="all")
    return df

def convert_to_dataframe(uploaded_file):
    """
    Core converter:
    1) Baca bytes
    2) Jika terdeteksi HTML -> read_html
       else -> coba read_excel (xlrd)
    3) Hapus 8 baris pertama (untuk HTML manual, untuk Excel via skiprows)
    4) Bersihkan DF
    """
    content = uploaded_file.read()
    uploaded_file.seek(0)

    # Coba HTML dulu jika file "beraroma" HTML
    if is_probably_html(content):
        try:
            df, err = try_read_html_to_df(content)
            if err:
                return None, err
            # Hapus 8 baris pertama secara manual (karena read_html tidak punya skiprows)
            df = df.iloc[8:] if len(df) > 8 else df.iloc[0:0]
            df = clean_df_after_skip(df)
            return df, None
        except Exception as e_html:
            # Jika HTML gagal, coba paksa sebagai Excel
            try:
                df = try_read_xls_to_df(content, skiprows=8)
                df = clean_df_after_skip(df)
                return df, None
            except Exception as e_xls:
                return None, f"Tidak bisa baca sebagai HTML maupun XLS. HTML err: {e_html}; XLS err: {e_xls}"

    # Jika tidak terdeteksi HTML, coba sebagai XLS
    try:
        df = try_read_xls_to_df(content, skiprows=8)
        df = clean_df_after_skip(df)
        return df, None
    except Exception as e_xls:
        # Fallback terakhir: coba HTML (kadang header HTML tidak muncul di 4KB awal)
        try:
            df, err = try_read_html_to_df(content)
            if err:
                return None, err
            df = df.iloc[8:] if len(df) > 8 else df.iloc[0:0]
            df = clean_df_after_skip(df)
            return df, None
        except Exception as e_html:
            return None, f"Tidak bisa baca file sebagai XLS maupun HTML. XLS err: {e_xls}; HTML err: {e_html}"

# -----------------------------
# Main action
# -----------------------------
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

                    # Buat nama file baru
                    new_filename = os.path.splitext(uploaded_file.name)[0] + ".xlsx"
                    output_path = os.path.join(tmpdir, new_filename)

                    try:
                        df.to_excel(output_path, index=False, engine='openpyxl')
                        zipf.write(output_path, arcname=new_filename)
                        success_count += 1
                        st.success(f"✅ Berhasil: {new_filename}")
                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan {new_filename}: {e}")

            # Jika ada minimal satu file berhasil, tampilkan tombol unduh
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
