import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Aplikasi Input & Export Excel", layout="wide")

st.title("📦 Aplikasi Packing - Input dan Download Data")

# --- 1. Upload File Excel Awal ---
uploaded_file = st.file_uploader("Unggah File Excel Awal", type=["xlsx"])

if uploaded_file:
    # Gunakan session_state agar data tidak hilang setelah interaksi
    if "sheets_dict" not in st.session_state:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        st.session_state.sheets_dict = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in sheet_names}
        st.session_state.sheet_names = sheet_names

    sheet_names = st.session_state.sheet_names
    sheets_dict = st.session_state.sheets_dict

    st.success(f"✅ File berhasil dimuat! Ditemukan {len(sheet_names)} sheet.")

    # Pilih sheet yang ingin diedit
    selected_sheet = st.selectbox("Pilih Sheet yang ingin ditampilkan dan diedit:", sheet_names)

    # Ambil sheet terpilih dari session_state
    df = sheets_dict[selected_sheet]

    st.subheader(f"📊 Data Awal dari Sheet: {selected_sheet}")
    st.dataframe(df, use_container_width=True)

    # --- 2. Input Data Baru ---
    st.subheader("📝 Tambahkan Data Baru ke Sheet yang Dipilih")
    with st.form("input_form", clear_on_submit=True):
        new_data = {}
        for col in df.columns:
            new_data[col] = st.text_input(f"{col}", "")
        submitted = st.form_submit_button("Tambah Data")

        if submitted:
            # Tambahkan data baru ke bagian atas sheet terpilih
            updated_df = pd.concat([pd.DataFrame([new_data]), df], ignore_index=True)
            st.session_state.sheets_dict[selected_sheet] = updated_df  # simpan update
            st.success(f"✅ Data baru berhasil ditambahkan ke atas sheet '{selected_sheet}'!")
            st.dataframe(updated_df, use_container_width=True)

    # --- 3. Download Excel (Semua Sheet Termasuk yang Diedit) ---
    st.subheader("⬇️ Unduh Hasil (Semua Sheet Termasuk yang Diedit)")

    def convert_all_sheets_to_excel(updated_sheets):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, data in updated_sheets.items():
                data.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()

    excel_data = convert_all_sheets_to_excel(st.session_state.sheets_dict)

    st.download_button(
        label="💾 Download Excel (Semua Sheet Terupdate)",
        data=excel_data,
        file_name="hasil_update_semua_sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Silakan unggah file Excel terlebih dahulu untuk memulai.")
