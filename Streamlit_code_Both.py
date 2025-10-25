# -*- coding: utf-8 -*-
"""
Created on Sun Oct 26 00:44:04 2025

@author: mreem
"""

import streamlit as st
import tempfile
import os
import base64
import Main_Code_Task
import Delta_code_5G

# ============================================================
# ---- PAGE CONFIG ----
# ============================================================
st.set_page_config(
    page_title="Network KPI PowerPoint Updater",
    page_icon="📊",
    layout="centered"
)

# ============================================================
# ---- ANIMATED BACKGROUND ----
# ============================================================
def add_animated_bg(image_file):
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
            animation: moveBg 40s ease-in-out infinite alternate;
        }}

        @keyframes moveBg {{
            0% {{ background-position: center top; }}
            100% {{ background-position: center bottom; }}
        }}

        /* Title styling */
        h1 {{
            color: #0f172a;
            text-shadow: 2px 2px 8px rgba(255,255,255,0.9);
        }}

        /* Section headers */
        .stHeader, h2, h3 {{
            color: #1e3a8a !important;
            font-weight: 800 !important;
            text-shadow: 1px 1px 3px rgba(255,255,255,0.7);
        }}

        /* Card style */
        .stCard {{
            background-color: rgba(255,255,255,0.85);
            border-radius: 16px;
            padding: 1.5rem;
            box-shadow: 0 6px 25px rgba(0,0,0,0.2);
        }}

        /* Buttons */
        .stButton>button {{
            background: linear-gradient(90deg, #2563eb, #60a5fa);
            color: white !important;
            border: none;
            border-radius: 10px;
            padding: 0.6rem 1.2rem;
            font-weight: 600;
            transition: 0.3s;
        }}
        .stButton>button:hover {{
            background: linear-gradient(90deg, #1d4ed8, #3b82f6);
            transform: scale(1.05);
        }}

        /* Radio buttons */
        .stRadio label {{
            color: #0f172a !important;
            font-weight: 600;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Apply the moving wallpaper
add_animated_bg("Containers_Angled_Amplifier_16x9.jpg")

# ============================================================
# ---- HEADER ----
# ============================================================
st.title("📊 Network KPI Weekly Slides Generator")
st.markdown("""
<div style='font-size:1.1rem; font-weight:500; color:#0f172a; background-color: rgba(255,255,255,0.7); 
padding:10px; border-radius:8px;'>
Select your report type (<b>UE & SI</b> or <b>DE</b>), upload the required Excel and PowerPoint files,  
and click <b>Run Processing</b> to automatically update your PowerPoint report.
</div>
""", unsafe_allow_html=True)

# ============================================================
# ---- REPORT TYPE ----
# ============================================================
report_type = st.radio("Select Report Type:", ["UE & SI", "DE"], horizontal=True)

# ============================================================
# ---- UE & SI SECTION ----
# ============================================================
if report_type == "UE & SI":
    st.markdown("<h2>📁 UE & SI Input Files</h2>", unsafe_allow_html=True)

    with st.container():
        st.markdown("<div class='stCard'>", unsafe_allow_html=True)

        excel_file = st.file_uploader("📈 Upload Excel file (.xlsx)", type=["xlsx"])
        ppt_file = st.file_uploader("📊 Upload PowerPoint file (.pptx)", type=["pptx"])

        if not (excel_file and ppt_file):
            st.info("Please upload both an Excel file and a PowerPoint file to continue.")
            st.stop()

        # ---- Save to temp ----
        temp_dir = tempfile.mkdtemp()
        excel_path = os.path.join(temp_dir, excel_file.name)
        pptx_path = os.path.join(temp_dir, ppt_file.name)

        with open(excel_path, "wb") as f:
            f.write(excel_file.read())
        with open(pptx_path, "wb") as f:
            f.write(ppt_file.read())

        st.success("✅ Files uploaded and saved temporarily.")

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
        st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# ---- DE SECTION ----
# ============================================================
else:
    st.markdown("<h2>📁 DE Input Files</h2>", unsafe_allow_html=True)

    with st.container():
        st.markdown("<div class='stCard'>", unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            excel_file_2G_3G_4G = st.file_uploader("📶 Upload 2G / 3G / 4G Excel file (.xlsx)", type=["xlsx"])
        with col2:
            excel_file_5G = st.file_uploader("📡 Upload 5G Excel file (.xlsx)", type=["xlsx"])

        ppt_file = st.file_uploader("📊 Upload PowerPoint file (.pptx)", type=["pptx"])

        if not (excel_file_2G_3G_4G and excel_file_5G and ppt_file):
            st.info("Please upload both Excel files (2G/3G/4G and 5G) and the PowerPoint file.")
            st.stop()

        # ---- Save to temp ----
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

        if st.button("🚀 Run Processing"):
            with st.spinner("Processing DE Report — please wait..."):
                try:
                    if hasattr(Delta_code_5G, 'main_with_paths_DE'):
                        Delta_code_5G.main_with_paths_DE(
                            excel_path_2G_3G_4G,
                            excel_path_5G,
                            pptx_path
                        )
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

        st.markdown("</div>", unsafe_allow_html=True)
