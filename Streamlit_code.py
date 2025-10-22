import streamlit as st
import tempfile
import os
import Main_Code_Task  # your full script file in repo root

# ---- Page Config ----
st.set_page_config(page_title="Network KPI PowerPoint Updater", page_icon="📊", layout="centered")

# ---- Background Image Setup ----
def add_bg_from_local(image_file):
    with open(image_file, "rb") as f:
        base64_image = f.read()
    import base64
    base64_image = base64.b64encode(base64_image).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{base64_image}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Example: use your own image path here
add_bg_from_local("background.jpg")  # <-- put your image file in the same folder as app.py

# ---- App Title ----
st.title("📊 UE & SI Customer KPIs Weekly Slides")
st.markdown("""
Upload your *Excel KPI file (.xlsx) "Orange_Agreed_KPIs"* and *PowerPoint template (.pptx) "Last Updated one"*,  
then click *Run Processing* to update the report automatically.
""")

# ---- File Uploads ----
excel_file = st.file_uploader("📈 Upload Excel file (.xlsx)", type=["xlsx"])
ppt_file = st.file_uploader("📊 Upload PowerPoint file (.pptx)", type=["pptx"])

if not (excel_file and ppt_file):
    st.info("Please upload both an Excel file and a PowerPoint file.")
    st.stop()

# ---- Temporary Save ----
temp_dir = tempfile.mkdtemp()
excel_path = os.path.join(temp_dir, excel_file.name)
pptx_path = os.path.join(temp_dir, ppt_file.name)

with open(excel_path, "wb") as f:
    f.write(excel_file.read())
with open(pptx_path, "wb") as f:
    f.write(ppt_file.read())

st.success("✅ Files saved to temporary folder.")

# ---- Checkbox ----
use_win32_if_available = st.checkbox("Allow Windows COM steps if running on Windows (ignored on Linux)", value=False)

# ---- Run Processing ----
if st.button("🚀 Run Processing"):
    with st.spinner("Processing — this may take a little while..."):
        try:
            Main_Code_Task.main_with_paths(excel_path, pptx_path)
            if hasattr(Main_Code_Task, 'main'):
                Main_Code_Task.main()
            else:
                raise RuntimeError("Main_Code_Task.py does not expose a main() function.")

            st.success("🎉 PowerPoint updated successfully!")

            with open(pptx_path, "rb") as f:
                st.download_button("⬇️ Download Updated PowerPoint", f, file_name="Updated_Report.pptx")

        except Exception as e:
            st.error(f"❌ Processing failed: {e}")
            st.exception(e)
