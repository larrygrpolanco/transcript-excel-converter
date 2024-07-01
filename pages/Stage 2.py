import streamlit as st
import pandas as pd


# Function to apply template to raw data
def apply_template(raw_data, template):
    merged_data = template.copy()
    merged_data["Speaker"] = raw_data["Speaker"]
    merged_data["Teacher (T) or Child (C)"] = raw_data["Teacher (T) or Child (C)"]
    merged_data["Utterance/Idea Units"] = raw_data["Utterance/Idea Units"]
    return merged_data


# Stage 2 Page
def main():
    st.title("Stage 2: Apply Template to Raw Excel")

    st.markdown(
        """
    ### Instructions:
    1. Upload the raw Excel file.
    2. Upload the Excel template file.
    3. The merged Excel file with the template will be available for download.
    """
    )

    raw_file = st.file_uploader("Upload Raw Excel File", type="xlsx")
    template_file = st.file_uploader("Upload Template File", type="xlsx")

    if raw_file is not None and template_file is not None:
        raw_data = pd.read_excel(raw_file)
        template = pd.read_excel(template_file)

        merged_data = apply_template(raw_data, template)

        st.markdown("### Merged Data with Template:")
        st.write(merged_data)

        if st.button("Download Final Excel"):
            merged_data.to_excel("final_transcript.xlsx", index=False)
            st.download_button(
                label="Download Final Excel File",
                data=open("final_transcript.xlsx", "rb").read(),
                file_name="final_transcript.xlsx",
            )


if __name__ == "__main__":
    main()
