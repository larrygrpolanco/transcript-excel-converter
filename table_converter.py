import os
import re
import pandas as pd
from docx import Document


def read_docx_to_text(file_path):
    """Reads DOCX file and returns all text as a single string."""
    doc = Document(file_path)
    full_text = "\n".join(paragraph.text for paragraph in doc.paragraphs)
    return full_text


def classify_speaker(speaker_id):
    """Classifies speaker as Teacher (T) or Child (C) based on ID."""
    if speaker_id.startswith("4"):
        return "C"
    elif speaker_id.startswith("3"):
        return "T"
    else:
        return None  # Unknown or misformatted ID


def parse_text(text):
    """Parses plain text and returns structured data, handling multi-line utterances."""
    data = []
    regex_pattern = r"^(\d+)\s*[:\-]?\s*(.*)"
    current_speaker_id = None
    accumulated_text = ""

    # Split the text into lines to simulate the original list of paragraphs
    text_list = text.split("\n")

    for line in text_list:
        match = re.match(regex_pattern, line)
        if match:
            if accumulated_text:  # Process the accumulated utterance before resetting
                speaker_type = classify_speaker(current_speaker_id)
                if speaker_type:
                    data.append(
                        {
                            "Speaker": current_speaker_id,
                            "Teacher (T) or Child (C)": speaker_type,
                            "Utterance/Idea Units": accumulated_text.strip(),
                        }
                    )
            current_speaker_id = match.group(1)
            accumulated_text = match.group(
                2
            ).strip()  # Start accumulating text for the new speaker
        else:
            if accumulated_text:  # Continue accumulating text if already started
                accumulated_text += " " + line.strip()

    # Don't forget to add the last accumulated utterance
    if accumulated_text:
        speaker_type = classify_speaker(current_speaker_id)
        if speaker_type:
            data.append(
                {
                    "Speaker": current_speaker_id,
                    "Teacher (T) or Child (C)": speaker_type,
                    "Utterance/Idea Units": accumulated_text.strip(),
                }
            )

    return data


def process_folder(input_folder, output_folder):
    """Processes all DOCX files in the input folder, outputting Excel files in the output folder."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(input_folder):
        if file_name.endswith(".docx"):
            file_path = os.path.join(input_folder, file_name)
            text = read_docx_to_text(file_path)  # Adjusted to read as plain text
            structured_data = parse_text(text)

            df = pd.DataFrame(structured_data)
            excel_file_name = file_name.replace(".docx", ".xlsx")
            excel_file_path = os.path.join(output_folder, excel_file_name)
            df.to_excel(excel_file_path, index=False)

            print(f"Processed {file_name} to {excel_file_path}")


def main():
    input_folder = "All Transcripts"  # "path/to/input/folder"
    output_folder = "Modified Transcripts"  # "path/to/output/folder"
    process_folder(input_folder, output_folder)


if __name__ == "__main__":
    main()
