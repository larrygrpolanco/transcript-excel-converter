import streamlit as st
import pandas as pd
import xlwings as xw
import tempfile
import os


def read_excel(file):
    return pd.read_excel(file)


def apply_template_with_formulas_and_validation(
    raw_data, template_file_path, output_file_path
):
    try:
        with xw.App(visible=False) as app:
            template_wb = app.books.open(template_file_path)
            template_ws = template_wb.sheets[0]

            # Convert raw_data DataFrame to list of lists
            raw_data_list = raw_data.values.tolist()

            # Write raw_data to the template sheet, starting from the second row to preserve headers/formulas
            for i, row in enumerate(raw_data_list, start=2):
                template_ws.range(f"A{i}").value = row

            # Save the updated workbook to the output file path
            template_wb.save(output_file_path)
            template_wb.close()

        return output_file_path
    except Exception as e:
        st.error(f"Error during applying template: {e}")
        return None


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
        st.write("Template uploaded successfully!")

        st.subheader("Upload Raw Excel File")
        raw_file = st.file_uploader("Upload Raw Excel File", type="xlsx")

        if raw_file:
            raw_data = read_excel(raw_file)
            st.write(f"Raw Excel file '{raw_file.name}' uploaded successfully!")

            with st.expander("Preview Raw Excel File"):
                st.write(raw_data)

            output_path = "output.xlsx"

            if st.button("Apply Template"):
                try:
                    # Save the template file to a temporary path
                    with tempfile.NamedTemporaryFile(
                        delete=False, suffix=".xlsx"
                    ) as temp_template_file:
                        temp_template_file.write(template_file.read())
                        temp_template_file_path = temp_template_file.name

                    applied_data_output_path = (
                        apply_template_with_formulas_and_validation(
                            raw_data, temp_template_file_path, output_path
                        )
                    )

                    if applied_data_output_path:
                        with open(applied_data_output_path, "rb") as f:
                            st.download_button(
                                label="Download Excel File with Applied Template",
                                data=f.read(),
                                file_name=f"applied_{raw_file.name}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                except Exception as e:
                    st.error(f"Error applying template: {e}")
                finally:
                    # Clean up temporary files
                    if os.path.exists(temp_template_file_path):
                        os.remove(temp_template_file_path)


if __name__ == "__main__":
    main()
