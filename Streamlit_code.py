# app.py
import streamlit as st
import tempfile
import os
import importlib
import Main_Code_Task  # your full script file in repo root

st.set_page_config(page_title="Network KPI PowerPoint Updater", page_icon="📊", layout="centered")
st.title("📊 Custromer KPIs Weekly Slides")
st.markdown("""
Upload your *Excel KPI file (.xlsx)* and *PowerPoint template (.pptx)*,  
then click *Run Processing* to update the report automatically.
""")

excel_file = st.file_uploader("📈 Upload Excel file (.xlsx)", type=["xlsx"])
ppt_file = st.file_uploader("📊 Upload PowerPoint file (.pptx)", type=["pptx"])

if not (excel_file and ppt_file):
    st.info("Please upload both an Excel file and a PowerPoint file.")
    st.stop()

# Save uploaded files to a temporary working directory
temp_dir = tempfile.mkdtemp()
excel_path = os.path.join(temp_dir, excel_file.name)
pptx_path = os.path.join(temp_dir, ppt_file.name)

with open(excel_path, "wb") as f:
    f.write(excel_file.read())
with open(pptx_path, "wb") as f:
    f.write(ppt_file.read())

st.success("✅ Files saved to temporary folder.")

# Provide a small checkbox to control whether to run windows-only COM steps
use_win32_if_available = st.checkbox("Allow Windows COM steps if running on Windows (ignored on Linux)", value=False)

if st.button("🚀 Run Processing"):
    with st.spinner("Processing — this may take a little while..."):
        try:
            # Inject the uploaded paths into the Main_Code_Task module globals if those global names are used.
            # Many scripts define global variables like 'excel_path' and 'pptx_file' — override them if present.
            Main_Code_Task._dict_['excel_path'] = excel_path
            Main_Code_Task._dict_['pptx_file'] = pptx_path



            # If your script exposes a callable main() run it; otherwise attempt to import/execute.
            if hasattr(Main_Code_Task, 'main'):
                Main_Code_Task.main()
            else:
                # fallback: try to run top-level function name or raise
                raise RuntimeError("Main_Code_Task.py does not expose a main() function.")

            st.success("🎉 PowerPoint updated successfully!")

            # Offer the user the updated pptx to download
            with open(pptx_path, "rb") as f:
                st.download_button("⬇️ Download Updated PowerPoint", f, file_name="Updated_Report.pptx")
        except Exception as e:
            st.error(f"❌ Processing failed: {e}")
            st.exception(e)



