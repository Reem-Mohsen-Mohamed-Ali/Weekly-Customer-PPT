# -*- coding: utf-8 -*-
"""
Created on Sun Oct 26 00:44:04 2025

@author: mreem
"""

import streamlit as st
import tempfile
import os
import Main_Code_Task
import Delta_code_5G  # your backend Python file

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
        /* --- Fullscreen Background --- */
        .stApp {{
            background-image: url("data:image/png;base64,{base64_image}");
            background-size: cover;
            background-repeat: no-repeat;
            background-position: center;
            background-attachment: fixed;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }}

        /* --- Glass Effect Box --- */
        .report-box {{
            background: rgba(255, 255, 255, 0.85);
            padding: 2.5rem 3rem;
            border-radius: 25px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3);
            backdrop-filter: blur(10px);
            max-width: 850px;
            width: 100%;
            text-align: center;
        }}

        /* --- Title --- */
        h1 {{
            text-align: center;
            font-weight: 900;
            color: #002b5c;
            text-shadow: 1px 1px 3px rgba(0,0,0,0.2);
            margin-bottom: 0.5rem;
        }}

        /* --- Subtitle --- */
        .subtitle {{
            text-align: center;
            font-size: 1.1rem;
            color: #003366;
            margin-bottom: 2rem;
        }}

        /* --- Radio Buttons --- */
        div.row-widget.stRadio > div {{
            justify-content: center;
            display: flex;
        }}

        label[data-testid="stMarkdownContainer"] p {{
            text-align: center !important;
        }}

        /* --- Button Styling --- */
        div.stButton > button:first-child {{
            background-color: #0073e6;
            color: white;
            font-size: 18px;
            border-radius: 12px;
            height: 3rem;
            width: 80%;
            margin: 1rem auto;
            border: none;
            box-shadow: 0 4px 10px rgba(0,0,0,0.3);
            transition: all 0.3s ease;
            display: block;
        }}
        div.stButton > button:first-child:hover {{
            background-color: #005bb5;
            transform: scale(1.03);
        }}

        /* --- Upload Elements --- */
        section[data-testid="stFileUploader"] {{
            text-align: center;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Add your background
add_bg_from_local("Containers_Angled_Amplifier_16x9.jpg")

# ---- Main App Box ----
st.markdown('<div class="report-box">', unsafe_allow_html=True)

st.title("📊 Network KPI Weekly Slides Generator")
st.markdown(
    '<p class="subtitle">Select your report type (<b>UE & SI</b> or <b>DE</b>), upload the required Excel and PowerPoint files,<br>then click <b>Run Processing</b> to automatically update your PowerPoint report.</p>',
    unsafe_allow_html=True
)

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
    else:
        temp_dir = tempfile.mkdtemp()
        excel_path = os.path.join(temp_dir, excel_file.name)
        pptx_path = os.path.join(temp_dir, ppt_file.name)

        with open(excel_path, "wb") as f:
            f.write(excel_file.read())
        with open(pptx_path, "wb") as f:
            f.write(ppt_file.read())

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
    else:
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

        if st.button("🚀 Run Processing"):
            with st.spinner("Processing DE Report — please wait..."):
                try:
                    if hasattr(Delta_code_5G, 'main_with_paths_DE'):
                        Delta_code_5G.main_with_paths_DE(excel_path_2G_3G_4G, excel_path_5G, pptx_path)
                    else:
                        Delta_code_5G.main_with_paths(excel_path_2G_3G_4G, pptx_path)
                    if hasattr(Delta_code_5G, 'main'):
                        Delta_code_5G.main()
                    st.success("🎉 DE PowerPoint updated successfully!")
                    with open(pptx_path, "rb") as f:
                        st.download_button("⬇️ Download Updated PowerPoint", f, file_name="Updated_DE_Report.pptx")
                except Exception as e:
                    st.error(f"❌ Processing failed: {e}")
                    st.exception(e)

st.markdown('</div>', unsafe_allow_html=True)
