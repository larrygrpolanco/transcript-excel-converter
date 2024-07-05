# Excel Template App

Excel Template App is a Python-based desktop application designed to streamline the process of converting transcript files to Excel format and applying predefined templates. This tool is particularly useful for research assistants working with transcript data.

## Features

- Convert .docx transcript files to raw Excel format
- Apply Excel templates to processed data
- Batch processing capabilities
- User-friendly graphical interface

## Installation

### Prerequisites

- Python 3.6 or higher
- pip (Python package installer)

### Steps

1. Open a command prompt or terminal.
2. Run the following command to install the Excel Template App:

   ```
   pip install excel_template_app
   ```

## Usage

After installation, you can run the application using one of the following methods:

1. From the command line:
   ```
   excel_template_app
   ```

2. From Python:
   ```python
   from excel_template_app import ExcelTemplateApp
   import tkinter as tk

   root = tk.Tk()
   app = ExcelTemplateApp(root)
   root.mainloop()
   ```

## Application Workflow

1. **Stage 1: Transcript to Raw Excel**
   - Select multiple .docx transcript files
   - Process the transcripts to convert them to raw Excel format
   - View processing results in the application window

2. **Stage 2: Apply Template**
   - Select an Excel template file
   - Select multiple raw Excel files generated from Stage 1
   - Apply the template to all selected files
   - Choose a directory to save the processed files

## Troubleshooting

- If you encounter any issues, ensure you have the latest version of the app:
  ```
  pip install --upgrade excel_template_app
  ```
- Check that all dependencies are correctly installed
- For persistent problems, please contact [Your Contact Information]

## Contributing

If you'd like to contribute to the project, please contact me

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contact

Larry Grullon-Polanco - larrygrpolanco@gmail.com

Project Link: https://github.com/larrygrpolanco/transcript-excel-converter
