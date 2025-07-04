
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel VLOOKUP Tool", page_icon="📊", layout="centered")

st.title("📊 Excel VLOOKUP Tool")

st.markdown(
    """
    อัปโหลด **F1** (ข้อมูลหลัก) และ **F2** (ตารางที่ใช้เทียบ)
    โปรแกรมจะจับคู่ค่าที่ตรงกันในคอลัมน์คีย์ (เช่น `PID`) แล้วดึงคอลัมน์ที่คุณเลือกจาก F2 มาต่อใน F1
    """
)

# ฟังก์ชันอ่านไฟล์ Excel
def read_excel(file):
    try:
        sheets = pd.read_excel(file, sheet_name=None)
        if isinstance(sheets, dict):
            df = pd.concat(sheets.values(), ignore_index=True)
        else:
            df = sheets
        return df
    except Exception as e:
        st.error(f"ไม่สามารถอ่านไฟล์: {e}")
        return pd.DataFrame()

f1_file = st.file_uploader("📤 อัปโหลด F1 (ข้อมูลหลัก)", type=["xlsx", "xls"], key="f1")
f2_file = st.file_uploader("📤 อัปโหลด F2 (ตารางอ้างอิง)", type=["xlsx", "xls"], key="f2")

if f1_file and f2_file:
    df1 = read_excel(f1_file)
    df2 = read_excel(f2_file)

    if df1.empty or df2.empty:
        st.stop()

    st.subheader("🔧 ตั้งค่าการจับคู่")
    pid_col1 = st.selectbox("คอลัมน์คีย์ใน F1", df1.columns.tolist(), index=df1.columns.tolist().index("PID") if "PID" in df1.columns else 0)
    pid_col2 = st.selectbox("คอลัมน์คีย์ใน F2", df2.columns.tolist(), index=df2.columns.tolist().index("PID") if "PID" in df2.columns else 0)
    result_col = st.selectbox("คอลัมน์ผลลัพธ์จาก F2", [c for c in df2.columns if c != pid_col2])

    if st.button("🔄 ดึงข้อมูลและดาวน์โหลด"):
        merged = df1.merge(df2[[pid_col2, result_col]], left_on=pid_col1, right_on=pid_col2, how="left")
        merged.rename(columns={result_col: f"{result_col}_lookup"}, inplace=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Result")

        st.success("เสร็จเรียบร้อย! ดาวน์โหลดไฟล์ได้เลย ↓")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ผลลัพธ์",
            data=buffer.getvalue(),
            file_name="vlookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("🔍 ดูตัวอย่าง F1"):
        st.dataframe(df1.head())

    with st.expander("🔍 ดูตัวอย่าง F2"):
        st.dataframe(df2.head())

else:
    st.info("กรุณาอัปโหลดไฟล์ทั้งสองก่อนเริ่มทำงาน")
