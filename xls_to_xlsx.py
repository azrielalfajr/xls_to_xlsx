import streamlit as st
import pandas as pd
import io
import zipfile
import os
import html5lib
from tempfile import TemporaryDirectory

st.set_page_config(page_title="XLS to XLSX Converter", layout="centered")
st.title("📄 Konversi XLS ke XLSX dan Download ZIP")
st.write("Upload beberapa file .xls dan akan dikonversi ke .xlsx. Baris pertama sampai ke-8 akan otomatis dihapus. Semua file hasil akan diunduh dalam bentuk .zip.")

uploaded_files = st.file_uploader("Upload file .xls", type=["xls"], accept_multiple_files=True)

def convert_to_dataframe(uploaded_file):
    content = uploaded_file.read()
    uploaded_file.seek(0)

    try:
        # Coba baca sebagai Excel dan skip 8 baris pertama
        df = pd.read_excel(io.BytesIO(content), engine="xlrd", skiprows=8)
        return df, None
    except Exception as e_excel:
        try:
            # Jika gagal, coba baca sebagai HTML
            dfs = pd.read_html(io.BytesIO(content))
            if dfs:
                df = dfs[0]
                # Hapus 8 baris pertama secara manual
                df = df.iloc[8:].reset_index(drop=True)
                return df, None
            else:
                return None, "Tidak ditemukan tabel dalam file HTML."
        except Exception as e_html:
            return None, f"Tidak bisa baca file: {e_html}"

if uploaded_files:
    if st.button("🔄 Konversi dan Download ZIP"):
        with TemporaryDirectory() as tmpdir:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

                for uploaded_file in uploaded_files:
                    df, error = convert_to_dataframe(uploaded_file)

                    if error:
                        st.error(f"Gagal mengonversi file {uploaded_file.name}: {error}")
                        continue

                    # Buat nama file baru
                    new_filename = os.path.splitext(uploaded_file.name)[0] + ".xlsx"
                    output_path = os.path.join(tmpdir, new_filename)

                    try:
                        df.to_excel(output_path, index=False, engine='openpyxl')
                        zipf.write(output_path, arcname=new_filename)
                    except Exception as e:
                        st.error(f"Gagal menyimpan file {new_filename}: {e}")

            zip_buffer.seek(0)
            st.success("✅ Konversi selesai!")
            st.download_button(
                label="📥 Download hasil (.zip)",
                data=zip_buffer,
                file_name="converted_xlsx_files.zip",
                mime="application/zip"
            )
