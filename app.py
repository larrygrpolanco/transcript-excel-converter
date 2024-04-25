import os
import re
import pandas as pd
from docx import Document
import streamlit as st
from PIL import Image

# Set page title and icon
st.set_page_config(page_title="Transcript Converter", page_icon=":file_folder:")

# Load logo
logo = Image.open("logo.png")  # Replace with your logo
st.sidebar.image(logo, use_column_width=True)

# Instructions
st.sidebar.header("Instructions")
st.sidebar.write("1. Upload one or multiple DOCX files.")
st.sidebar.write("2. Preview the conversion by clicking 'Preview'.")
st.sidebar.write("3. Download the converted Excel files by clicking 'Download'.")

# File uploader
st.sidebar.header("Upload Files")
uploaded_files = st.sidebar.file_uploader(
    "Select DOCX files", accept_multiple_files=True
)

# Preview button
preview_button = st.sidebar.button("Preview")

# Download button
download_button = st.sidebar.button("Download")

if uploaded_files:
    # Process files
    processed_files = []
    for file in uploaded_files:
        text = read_docx_to_text(file)
        structured_data = parse_text(text)
        df = pd.DataFrame(structured_data)
        processed_files.append(df)

    # Preview
    if preview_button:
        st.write("Preview:")
        st.dataframe(processed_files[0])  # Display the first file's preview

    # Download
    if download_button:
        for i, df in enumerate(processed_files):
            file_name = uploaded_files[i].name.replace(".docx", ".xlsx")
            df.to_excel(file_name, index=False)
            with open(file_name, "rb") as file:
                st.download_button("Download", file, file_name)
