import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="🔍 เครื่องมือรวมข้อมูล Excel By SURANAN ",
    page_icon="📊",
    layout="centered",
)

# ---------------------------------------------------------
# ส่วนหัวหน้าเว็บ
# ---------------------------------------------------------

st.title("🔍รวมข้อมูล Data Exchange V1")

st.markdown(
    """ **จัดทำโดย นาย สุรนันท์ ยามดี โรงพยาบาลส่งเสริมสุขภาพตำบลบ้านโนนค้อ** """
)

# ---------------------------------------------------------
# ฟังก์ชันอ่าน Excel
# ---------------------------------------------------------

def read_excel(uploaded_file: "st.uploaded_file") -> pd.DataFrame:
    """อ่านไฟล์ Excel ทุกชีตแล้วต่อกันเป็น DataFrame เดียว"""
    try:
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
        if isinstance(sheets, dict):
            return pd.concat(sheets.values(), ignore_index=True)
        return sheets
    except Exception as e:
        st.error(f"❌ ไม่สามารถอ่านไฟล์: {e}")
        return pd.DataFrame()

# ---------------------------------------------------------
# โซนอัปโหลดไฟล์
# ---------------------------------------------------------

f1_file = st.file_uploader("📤 อัปโหลด F1 (ข้อมูลจาก HDC)", type=["xlsx", "xls"], key="f1")
f2_file = st.file_uploader("📤 อัปโหลด F2 (ไฟล์อ้างอิงจาก HOSXP)", type=["xlsx", "xls"], key="f2")

if f1_file and f2_file:
    df1 = read_excel(f1_file)
    df2 = read_excel(f2_file)

    if df1.empty or df2.empty:
        st.stop()

    # -----------------------------------------------------
    # การตั้งค่าจับคู่และผลลัพธ์
    # -----------------------------------------------------

    st.subheader("⚙️ ตั้งค่าการจับคู่และผลลัพธ์")

    pid_col1 = st.selectbox(
        "คอลัมน์คีย์ใน F1 (สำหรับจับคู่)", df1.columns.tolist(), index=df1.columns.tolist().index("PID") if "PID" in df1.columns else 0
    )

    pid_col2 = st.selectbox(
        "คอลัมน์คีย์ใน F2 (สำหรับจับคู่)", df2.columns.tolist(), index=df2.columns.tolist().index("PID") if "PID" in df2.columns else 0
    )

    result_cols = st.multiselect(
        "เลือกคอลัมน์ผลลัพธ์จาก F2 (เลือกได้หลายคอลัมน์)",
        [c for c in df2.columns if c != pid_col2],
    )

    st.divider()

    # -----------------------------------------------------
    # ปุ่มประมวลผล
    # -----------------------------------------------------

    if st.button("🔄 ดึงข้อมูลและดาวน์โหลด"):
        if not result_cols:
            st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 คอลัมน์ผลลัพธ์ก่อนครับ")
            st.stop()

        # ทำการ merge ครั้งเดียวด้วยคอลัมน์ที่เลือกทั้งหมด
        merged = df1.merge(
            df2[[pid_col2] + result_cols],
            left_on=pid_col1,
            right_on=pid_col2,
            how="left",
        )

        # เปลี่ยนชื่อคอลัมน์ผลลัพธ์ให้มี _lookup ต่อท้าย (กันซ้ำ)
        rename_map = {col: f"{col}_lookup" for col in result_cols}
        merged.rename(columns=rename_map, inplace=True)

        # สร้างไฟล์ Excel ในหน่วยความจำแล้วให้ดาวน์โหลด
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Result")

        st.success("✅ ดึงข้อมูลเสร็จเรียบร้อย! ดาวน์โหลดได้เลย ↓")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ผลลัพธ์",
            data=buffer.getvalue(),
            file_name="vlookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -----------------------------------------------------
    # ตัวอย่างข้อมูล (optional)
    # -----------------------------------------------------

    with st.expander("🔍 ดูตัวอย่าง F1 (หัว 5 แถว)"):
        st.dataframe(df1.head())

    with st.expander("🔍 ดูตัวอย่าง F2 (หัว 5 แถว)"):
        st.dataframe(df2.head())

else:
    st.info("⬆️ กรุณาอัปโหลดไฟล์ F1 และ F2 ก่อนเริ่มใช้งานครับ")

# ---------------------------------------------------------
# แถบด้านข้าง (Sidebar)
# ---------------------------------------------------------

st.sidebar.header("📝 วิธีใช้งานฉบับย่อ")

st.sidebar.markdown(
    """
    **ข้อสำคัญ ทำข้อมูลที่จะเทียบต้องเหมือนกัน เอาเลข PID จะง่าย โดยใช้สูตร Excel ให้เลข PID ใน HosXp เหมือนกับ HDC คือทำให้เป็นเลข 6 หลัก =TEXT(ข้อมูล,"000000")
    และ SQL สำหรับดึงเลข PID จาก HosXp คือ SELECT * FROM person; คัดลอกไปใช้ได้เลยครับ** 
1. กด **Browse files** เพื่ออัปโหลดไฟล์ Excel F1 และ F2  
2. เลือกคอลัมน์คีย์ / คอลัมน์ผลลัพธ์
3. กด **ดึงข้อมูลและดาวน์โหลด**  
4. รับไฟล์ผลลัพธ์ `.xlsx` ในทันที  

**© Suranan Yamdi** 

ติดต่อสอบถาม elricscythe@gmail.com
"""
)
