import streamlit as st
import tempfile
import os
import Islam_Slides_Task  # your existing script

st.set_page_config(page_title="Network KPI PowerPoint Updater", page_icon="📊", layout="centered")

st.title("📊 Network KPI PowerPoint Updater")
st.markdown("""
Upload your *Excel KPI file (.xlsx)* and *PowerPoint template (.pptx)*,  
then click *Run Processing* to update the report automatically.
""")

excel_file = st.file_uploader("📈 Upload Excel file", type=["xlsx"])
ppt_file = st.file_uploader("📊 Upload PowerPoint file", type=["pptx"])

if excel_file and ppt_file:
    temp_dir = tempfile.mkdtemp()
    excel_path = os.path.join(temp_dir, excel_file.name)
    pptx_path = os.path.join(temp_dir, ppt_file.name)

    with open(excel_path, "wb") as f:
        f.write(excel_file.read())
    with open(pptx_path, "wb") as f:
        f.write(ppt_file.read())

    st.success("✅ Files uploaded successfully!")

    if st.button("🚀 Run Processing"):
        try:
            # override file paths dynamically
            Islam_Slides_Task.main._globals_['excel_path'] = excel_path
            Islam_Slides_Task.main._globals_['pptx_file'] = pptx_path

            with st.spinner("Processing and updating PowerPoint..."):
                Islam_Slides_Task.main()

            with open(pptx_path, "rb") as f:
                st.success("🎉 PowerPoint updated successfully!")
                st.download_button(
                    "⬇️ Download Updated PowerPoint",
                    f,
                    file_name="Updated_Report.pptx"
                )
        except Exception as e:
            st.error(f"❌ Error: {e}")
else:
    st.info("Please upload both files to continue.")