# -*- coding: utf-8 -*-
"""
Created on Sun Oct 26 00:44:04 2025

@author: mreem
"""

import streamlit as st
import tempfile
import os
import Main_Code_Task  # your backend Python file

# ---- Page Config ----
st.set_page_config(page_title="Network KPI PowerPoint Updater", page_icon="📊", layout="centered")

# ---- Background Image ----
def add_bg_from_local(image_file):
    import base64
    with open(image_file, "rb") as f:
        base64_image = base64.b64encode(f.read()).decode()
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

# Background image (place in same folder as this script)
add_bg_from_local("Containers_Angled_Amplifier_16x9.jpg")

# ---- App Header ----
st.title("📊 Network KPI Weekly Slides Generator")
st.markdown("""
Select your report type (**UE & SI** or **DE**), upload the required Excel and PowerPoint files,  
and click **Run Processing** to automatically update your PowerPoint report.
""")

# ---- Report Type Selection ----
report_type = st.radio("Select Report Type:", ["UE & SI", "DE"], horizontal=True)

# ============================================================
# ---- UE & SI SECTION ----
# ============================================================
if report_type == "UE & SI":
    st.header("📁 UE & SI Input Files")

    excel_file = st.file_uploader("📈 Upload Excel file (.xlsx)", type=["xlsx"])
    ppt_file = st.file_uploader("📊 Upload PowerPoint file (.pptx)", type=["pptx"])

    if not (excel_file and ppt_file):
        st.info("Please upload both an Excel file and a PowerPoint file to continue.")
        st.stop()

    # ---- Temporary Save ----
    temp_dir = tempfile.mkdtemp()
    excel_path = os.path.join(temp_dir, excel_file.name)
    pptx_path = os.path.join(temp_dir, ppt_file.name)

    with open(excel_path, "wb") as f:
        f.write(excel_file.read())
    with open(pptx_path, "wb") as f:
        f.write(ppt_file.read())

    st.success("✅ Files uploaded and saved temporarily.")

    # ---- Run Processing ----
    if st.button("🚀 Run Processing"):
        with st.spinner("Processing UE & SI Report — please wait..."):
            try:
                Main_Code_Task.main_with_paths(excel_path, pptx_path)
                if hasattr(Main_Code_Task, 'main'):
                    Main_Code_Task.main()

                st.success("🎉 UE & SI PowerPoint updated successfully!")
                with open(pptx_path, "rb") as f:
                    st.download_button("⬇️ Download Updated PowerPoint", f, file_name="Updated_UE_SI_Report.pptx")

            except Exception as e:
                st.error(f"❌ Processing failed: {e}")
                st.exception(e)

# ============================================================
# ---- DE SECTION ----
# ============================================================
else:
    st.header("📁 DE Input Files")

    excel_file_2G_3G_4G = st.file_uploader("📶 Upload 2G / 3G / 4G Excel file (.xlsx)", type=["xlsx"])
    excel_file_5G = st.file_uploader("📡 Upload 5G Excel file (.xlsx)", type=["xlsx"])
    ppt_file = st.file_uploader("📊 Upload PowerPoint file (.pptx)", type=["pptx"])

    if not (excel_file_2G_3G_4G and excel_file_5G and ppt_file):
        st.info("Please upload both Excel files (2G/3G/4G and 5G) and the PowerPoint file.")
        st.stop()

    # ---- Temporary Save ----
    temp_dir = tempfile.mkdtemp()
    excel_path_2G_3G_4G = os.path.join(temp_dir, excel_file_2G_3G_4G.name)
    excel_path_5G = os.path.join(temp_dir, excel_file_5G.name)
    pptx_path = os.path.join(temp_dir, ppt_file.name)

    for file_obj, path in [
        (excel_file_2G_3G_4G, excel_path_2G_3G_4G),
        (excel_file_5G, excel_path_5G),
        (ppt_file, pptx_path),
    ]:
        with open(path, "wb") as f:
            f.write(file_obj.read())

    st.success("✅ All DE files uploaded and saved temporarily.")

    # ---- Run Processing ----
    if st.button("🚀 Run Processing"):
        with st.spinner("Processing DE Report — please wait..."):
            try:
                # Expecting a DE-specific function in backend
                if hasattr(Main_Code_Task, 'main_with_paths_DE'):
                    Main_Code_Task.main_with_paths_DE(
                        excel_path_2G_3G_4G,
                        excel_path_5G,
                        pptx_path
                    )
                else:
                    # fallback if not separate
                    Main_Code_Task.main_with_paths(excel_path_2G_3G_4G, pptx_path)

                if hasattr(Main_Code_Task, 'main'):
                    Main_Code_Task.main()

                st.success("🎉 DE PowerPoint updated successfully!")
                with open(pptx_path, "rb") as f:
                    st.download_button("⬇️ Download Updated PowerPoint", f, file_name="Updated_DE_Report.pptx")

            except Exception as e:
                st.error(f"❌ Processing failed: {e}")
                st.exception(e)