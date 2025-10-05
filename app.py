import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Aplikasi Input & Export Excel", layout="wide")
st.title("📦 Aplikasi Packing - Input & Download Data Excel")

# ============================================================
# 🔹 Persiapan Data Template (Dari File Excel Acuan)
# ============================================================
TEMPLATE_PATH = "PACKING HLP 03 SEPTEMBER 2025 1.xlsx"

try:
    xls_template = pd.ExcelFile(TEMPLATE_PATH)
    template_sheets = {sheet: pd.read_excel(xls_template, sheet_name=sheet) for sheet in xls_template.sheet_names}
    st.sidebar.success(f"📘 Template ditemukan: {len(template_sheets)} sheet.")
except Exception as e:
    st.sidebar.error("⚠️ File template tidak ditemukan di folder aplikasi.")
    template_sheets = {}

# ============================================================
# 🔸 Pilihan Mode
# ============================================================
mode = st.radio("Pilih Mode:", ["📤 Upload File Excel", "📝 Input Data Baru"])

# ============================================================
# 🔹 MODE 1 — UPLOAD FILE EXCEL
# ============================================================
if mode == "📤 Upload File Excel":
    uploaded_file = st.file_uploader("Unggah File Excel Awal", type=["xlsx"])

    if uploaded_file:
        if "sheets_dict" not in st.session_state:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            # 🔄 Baca data dari bawah ke atas
            st.session_state.sheets_dict = {
                sheet: pd.read_excel(xls, sheet_name=sheet).iloc[::-1].reset_index(drop=True)
                for sheet in sheet_names
            }
            st.session_state.sheet_names = sheet_names

        sheet_names = st.session_state.sheet_names
        sheets_dict = st.session_state.sheets_dict

        st.success(f"✅ File berhasil dimuat! Ditemukan {len(sheet_names)} sheet.")

        selected_sheet = st.selectbox("Pilih Sheet yang ingin ditampilkan dan diedit:", sheet_names)

        df = sheets_dict[selected_sheet]
        st.subheader(f"📊 Data dari Sheet: {selected_sheet} (dibaca dari bawah ke atas)")
        st.dataframe(df, use_container_width=True)

        # Input data baru
        st.subheader("📝 Tambahkan Data Baru ke Sheet yang Dipilih")
        with st.form("input_form_upload", clear_on_submit=True):
            new_data = {col: st.text_input(f"{col}", "") for col in df.columns}
            submitted = st.form_submit_button("Tambah Data")

            if submitted:
                # Tambah data di atas (urutan logika dibalik)
                updated_df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
                updated_df = updated_df.iloc[::-1].reset_index(drop=True)
                st.session_state.sheets_dict[selected_sheet] = updated_df
                st.success(f"✅ Data baru berhasil ditambahkan ke atas sheet '{selected_sheet}'!")
                st.dataframe(updated_df, use_container_width=True)

        # Download hasil semua sheet
        st.subheader("⬇️ Unduh Hasil (Semua Sheet Termasuk yang Diedit)")

        def convert_all_sheets_to_excel(updated_sheets):
            output = BytesIO()
            non_empty_sheets = {k: v for k, v in updated_sheets.items() if not v.empty}
            if not non_empty_sheets:
                raise ValueError("Tidak ada sheet berisi data untuk disimpan!")
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet_name, data in non_empty_sheets.items():
                    # Simpan tetap dari bawah ke atas
                    data.iloc[::-1].to_excel(writer, index=False, sheet_name=sheet_name)
            return output.getvalue()

        try:
            excel_data = convert_all_sheets_to_excel(st.session_state.sheets_dict)
            st.download_button(
                label="💾 Download Excel (Semua Sheet Terupdate - Urutan Dibalik)",
                data=excel_data,
                file_name="hasil_update_semua_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except ValueError as e:
            st.warning(str(e))

    else:
        st.info("Silakan unggah file Excel terlebih dahulu untuk memulai.")

# ============================================================
# 🔹 MODE 2 — INPUT DATA BARU (DARI TEMPLATE EXCEL)
# ============================================================
else:
    if not template_sheets:
        st.error("❌ File template belum tersedia. Pastikan file Excel acuan ada di folder aplikasi.")
    else:
        st.subheader("🆕 Input Data Baru Berdasarkan Template")

        # Pilih sheet dari template
        selected_template = st.selectbox("Pilih Sheet Template:", list(template_sheets.keys()))
        df_template = template_sheets[selected_template]

        st.write(f"📄 Struktur kolom dari sheet **{selected_template}**:")
        st.dataframe(df_template.head(3), use_container_width=True)

        # Simpan state input manual
        if "input_sheets" not in st.session_state:
            st.session_state.input_sheets = {
                sheet: pd.DataFrame(columns=df.columns) for sheet, df in template_sheets.items()
            }

        df_input = st.session_state.input_sheets[selected_template]

        # Form input
        st.subheader(f"🧾 Tambahkan Data ke Sheet: {selected_template}")
        with st.form("input_form_template", clear_on_submit=True):
            new_data = {col: st.text_input(f"{col}", "") for col in df_template.columns}
            submitted = st.form_submit_button("Tambah Data")

            if submitted:
                df_input = pd.concat([df_input, pd.DataFrame([new_data])], ignore_index=True)
                df_input = df_input.iloc[::-1].reset_index(drop=True)
                st.session_state.input_sheets[selected_template] = df_input
                st.success(f"✅ Data baru berhasil ditambahkan ke sheet '{selected_template}' (dibaca dari bawah ke atas)!")
                st.dataframe(df_input, use_container_width=True)

        # Download hasil dua sheet (misal dua sheet yang sudah diinput)
        st.subheader("⬇️ Unduh Hasil Input Semua Sheet")

        def convert_input_sheets_to_excel(input_sheets):
            output = BytesIO()
            non_empty_sheets = {k: v for k, v in input_sheets.items() if not v.empty}
            if not non_empty_sheets:
                raise ValueError("Tidak ada sheet berisi data untuk disimpan!")
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet_name, df_data in non_empty_sheets.items():
                    df_data.iloc[::-1].to_excel(writer, index=False, sheet_name=sheet_name)
            return output.getvalue()

        try:
            excel_data = convert_input_sheets_to_excel(st.session_state.input_sheets)
            st.download_button(
                label="💾 Download Excel (Semua Sheet Input Baru - Urutan Dibalik)",
                data=excel_data,
                file_name="hasil_input_semua_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except ValueError as e:
            st.warning(str(e))
