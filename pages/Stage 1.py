import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO


# Function to process the transcript from a string
def process_transcript(transcript):
    data = {"Speaker": [], "Teacher (T) or Child (C)": [], "Utterance/Idea Units": []}
    lines = transcript.split("\n")
    for line in lines:
        if line.startswith("*"):
            parts = line.split(":")
            if len(parts) >= 2:
                speaker_id = parts[0][1:].strip()
                utterance = parts[1].strip()
                role = "T" if speaker_id.startswith("3") else "C"
                data["Speaker"].append(speaker_id)
                data["Teacher (T) or Child (C)"].append(role)
                data["Utterance/Idea Units"].append(utterance)
    return pd.DataFrame(data)


# Function to read .docx file and extract text
def read_docx(file):
    doc = Document(file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return "\n".join(full_text)


# Function to check transcript formatting
def is_correct_format(transcript):
    lines = transcript.split("\n")
    for line in lines:
        if line.strip() == "":  # Skip empty lines
            continue
        if not line.startswith("*") or ":" not in line:
            return False
        parts = line.split(":")
        if len(parts) < 2 or not parts[1].strip().endswith((".", "?", "!", ";")):
            return False
    return True


# Stage 1 Page
def main():
    st.title("Stage 1: Transcript to Raw Excel")

    st.markdown(
        """
    ### Instructions:
    1. Upload Word documents containing the transcripts.
    2. Ensure the file format is correct (.docx).
    3. The processed transcripts will be available for download as Excel files.
    """
    )

    uploaded_files = st.file_uploader(
        "Upload Transcript Files", type="docx", accept_multiple_files=True
    )

    if uploaded_files:
        conversion_data = {"File Name": [], "Successfully Converted": []}

        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            transcript = read_docx(uploaded_file)
            if is_correct_format(transcript):
                conversion_data["File Name"].append(file_name)
                conversion_data["Successfully Converted"].append("Yes")
            else:
                conversion_data["File Name"].append(file_name)
                conversion_data["Successfully Converted"].append("No")

        conversion_df = pd.DataFrame(conversion_data)

        st.subheader("Conversion Status")
        st.write(conversion_df)

        st.subheader("Download Processed Transcripts")
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name.split(".")[0]
            transcript = read_docx(uploaded_file)
            processed_data = process_transcript(transcript)

            with st.expander(f"{file_name} Preview"):
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.write(processed_data)
                with col2:
                    output = BytesIO()
                    processed_data.to_excel(output, index=False)
                    output.seek(0)

                    st.download_button(
                        label="Download",
                        data=output,
                        file_name=f"{file_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


if __name__ == "__main__":
    main()
