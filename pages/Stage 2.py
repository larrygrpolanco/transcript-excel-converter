import streamlit as st
import pandas as pd
from io import BytesIO


# Function to read Excel file and return DataFrame
def read_excel(file):
    return pd.read_excel(file)


# Function to merge raw data with template
def merge_with_template(raw_data, template):
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
    1. Upload the Excel template file first.
    2. Then upload a raw Excel file from Stage 1.
    3. Apply the template to the file and download it.
    """
    )

    st.subheader("Upload Template File")
    template_file = st.file_uploader("Upload Template File", type="xlsx")

    if template_file:
        template = read_excel(template_file)
        st.write("Template uploaded successfully!")
        st.write(template)

        st.subheader("Upload Raw Excel File")
        raw_file = st.file_uploader("Upload Raw Excel File", type="xlsx")

        if raw_file:
            raw_data = read_excel(raw_file)
            st.write(f"Raw Excel file '{raw_file.name}' uploaded successfully!")
            st.write(raw_data)

            merged_data = merge_with_template(raw_data, template)
            st.write("Merged Data:")
            st.write(merged_data)

            output = BytesIO()
            merged_data.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                label="Download Merged Excel File",
                data=output,
                file_name=f"merged_{raw_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Commented out batch downloading for troubleshooting
            # merged_files = []
            # for raw_file in raw_files:
            #     raw_data = read_excel(raw_file)
            #     merged_data = merge_with_template(raw_data, template)
            #     merged_files.append((raw_file.name, merged_data))

            # if merged_files:
            #     with zipfile.ZipFile("merged_files.zip", "w") as zf:
            #         for file_name, df in merged_files:
            #             output = BytesIO()
            #             df.to_excel(output, index=False)
            #             output.seek(0)
            #             zf.writestr(file_name, output.read())

            #     with open("merged_files.zip", "rb") as f:
            #         st.download_button(
            #             label="Download Merged Files",
            #             data=f,
            #             file_name="merged_files.zip",
            #             mime="application/zip"
            #         )


if __name__ == "__main__":
    main()
