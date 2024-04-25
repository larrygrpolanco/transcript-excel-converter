import pandas as pd
import os
import xlwings as xw

template_path = "Coding Template.xlsx"  # Path to the template file
raw_files_dir = (
    "Excel Raw Transcripts"  # Path to the directory containing the raw files
)
output_dir = "Final Transcripts"  # Path to the output directory

if not os.path.exists(output_dir):
    os.makedirs(output_dir)


def process_file(raw_file, template_path, output_path):
    # Open the template workbook
    app = xw.App(visible=False)  # Run Excel in the background
    wb = app.books.open(template_path)
    sheet = wb.sheets[0]  # Assuming data is in the first sheet

    # Load the raw data into a DataFrame
    raw_data = pd.read_excel(raw_file)

    # Start writing from a specific row, for example, row 2
    start_row = 2  # Adjust based on your template
    for idx, row in enumerate(raw_data.itertuples(index=False), start=start_row):
        # Correctly map columns to Excel
        sheet.range(f"A{idx}").value = row[0]  # Speaker ID
        sheet.range(f"B{idx}").value = row[1]  # Teacher or Child
        sheet.range(f"C{idx}").value = row[2]  # Utterance/Idea Units

    # Save and close
    wb.save(output_path)
    wb.close()
    app.quit()


# Process each file in the directory
for filename in os.listdir(raw_files_dir):
    if filename.endswith(".xlsx") and filename != os.path.basename(template_path):
        raw_file_path = os.path.join(raw_files_dir, filename)
        output_file_path = os.path.join(output_dir, f"{filename}")
        process_file(raw_file_path, template_path, output_file_path)
