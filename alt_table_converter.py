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
    data = []
    current_speaker_id = None
    accumulated_text = ""
    # Define known annotations (add to this list as necessary)
    known_annotations = {"PRE-ASSESSMENT", "POST-ASSESSMENT", "*RECORDING STARTED SHORTLY AFTER LESSON STARTED*", "ASSESSMENT"}

    text_list = text.split("\n")

    for line in text_list:
        line = line.strip()
        if not line:
            continue

        # Check if the line is a known annotation
        if line in known_annotations:
            if accumulated_text and current_speaker_id:
                # Save the current dialogue before the annotation
                speaker_type = classify_speaker(current_speaker_id)
                if speaker_type:
                    data.append(
                        {
                            "Speaker": current_speaker_id,
                            "Teacher (T) or Child (C)": speaker_type,
                            "Utterance/Idea Units": accumulated_text.strip(),
                        }
                    )
                accumulated_text = ""
            # Add the annotation entry
            data.append(
                {
                    "Speaker": "Annotation",
                    "Teacher (T) or Child (C)": "N/A",
                    "Utterance/Idea Units": line,
                }
            )
            current_speaker_id = (
                None  # Reset the current speaker ID after an annotation
            )
            continue

        if line.isdigit():
            if accumulated_text and current_speaker_id:
                # Save the current dialogue before starting new speaker ID
                speaker_type = classify_speaker(current_speaker_id)
                if speaker_type:
                    data.append(
                        {
                            "Speaker": current_speaker_id,
                            "Teacher (T) or Child (C)": speaker_type,
                            "Utterance/Idea Units": accumulated_text.strip(),
                        }
                    )
                accumulated_text = ""
            current_speaker_id = line
        else:
            # Continue accumulating text for the current speaker
            accumulated_text += " " + line if accumulated_text else line

    # Append last accumulated entry
    if accumulated_text and current_speaker_id:
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
    input_folder = "Problem Transcripts"  # "path/to/input/folder"
    output_folder = "Converted Transcript"  # "path/to/output/folder"
    process_folder(input_folder, output_folder)


if __name__ == "__main__":
    main()
