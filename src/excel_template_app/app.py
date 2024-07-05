import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import xlwings as xw
from docx import Document
import os


class ExcelTemplateApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel Template Applicator")
        master.geometry("700x500")

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.stage1_frame = ttk.Frame(self.notebook)
        self.stage2_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.stage1_frame, text="Stage 1: Transcript to Raw Excel")
        self.notebook.add(self.stage2_frame, text="Stage 2: Apply Template")

        self.setup_stage1()
        self.setup_stage2()

    def setup_stage1(self):
        instructions = (
            "Stage 1 Instructions:\n"
            "1. Click 'Select Files' to choose multiple .docx transcript files.\n"
            "2. Click 'Process Transcripts' to convert them to raw Excel files.\n"
            "3. Processing results will be displayed below."
        )
        ttk.Label(
            self.stage1_frame, text=instructions, wraplength=650, justify="left"
        ).pack(pady=10)

        self.transcript_files = []

        ttk.Button(
            self.stage1_frame, text="Select Files", command=self.load_transcripts
        ).pack(pady=5)
        ttk.Button(
            self.stage1_frame,
            text="Process Transcripts",
            command=self.process_transcripts,
        ).pack(pady=5)

        self.status_frame = ttk.Frame(self.stage1_frame)
        self.status_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.status_text = tk.Text(self.status_frame, wrap=tk.WORD, height=10, width=50)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(
            self.status_frame, orient="vertical", command=self.status_text.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_text.config(yscrollcommand=scrollbar.set)
        self.status_text.tag_configure("success", foreground="green")
        self.status_text.tag_configure("failure", foreground="red")

    def setup_stage2(self):
        instructions = (
            "Stage 2 Instructions:\n"
            "1. Select the template Excel file.\n"
            "2. Select multiple raw Excel files generated from Stage 1.\n"
            "3. Click 'Apply Template' to process all files.\n"
            "4. Choose a directory to save the processed files."
        )
        ttk.Label(
            self.stage2_frame, text=instructions, wraplength=650, justify="left"
        ).pack(pady=10)

        ttk.Label(self.stage2_frame, text="Template File:").pack(
            anchor="w", padx=5, pady=5
        )
        self.template_entry = ttk.Entry(self.stage2_frame, width=70)
        self.template_entry.pack(fill="x", padx=5, pady=2)
        ttk.Button(
            self.stage2_frame, text="Select Template", command=self.load_template
        ).pack(anchor="w", padx=5, pady=2)

        ttk.Label(self.stage2_frame, text="Raw Excel Files:").pack(
            anchor="w", padx=5, pady=5
        )
        self.raw_files_listbox = tk.Listbox(self.stage2_frame, width=70, height=5)
        self.raw_files_listbox.pack(fill="both", expand=True, padx=5, pady=2)
        ttk.Button(
            self.stage2_frame, text="Select Raw Files", command=self.load_raw_files
        ).pack(anchor="w", padx=5, pady=2)

        ttk.Button(
            self.stage2_frame,
            text="Apply Template",
            command=self.apply_template_to_multiple,
        ).pack(pady=10)

        self.stage2_status = tk.StringVar()
        ttk.Label(
            self.stage2_frame, textvariable=self.stage2_status, wraplength=650
        ).pack(pady=5)

    def load_template(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, file)

    def load_raw_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        for file in files:
            self.raw_files_listbox.insert(tk.END, file)

    def apply_template_to_multiple(self):
        template_file = self.template_entry.get()
        raw_files = list(self.raw_files_listbox.get(0, tk.END))

        if not template_file or not raw_files:
            messagebox.showerror("Error", "Please select both template and raw files.")
            return

        save_dir = filedialog.askdirectory(
            title="Select Directory to Save Processed Files"
        )
        if not save_dir:
            return

        success_count = 0
        failure_count = 0

        for raw_file in raw_files:
            try:
                raw_data = pd.read_excel(raw_file)
                base_name = os.path.basename(raw_file)
                save_path = os.path.join(save_dir, f"{base_name}")

                success = self.apply_template_with_direct_save(
                    raw_data, template_file, save_path
                )
                if success:
                    success_count += 1
                else:
                    failure_count += 1
            except Exception as e:
                messagebox.showerror(
                    "Error", f"An error occurred processing {base_name}: {str(e)}"
                )
                failure_count += 1

        self.stage2_status.set(
            f"Processing complete.\nSuccessful: {success_count}\nFailed: {failure_count}\nFiles saved in: {save_dir}"
        )

    def load_transcripts(self):
        files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")])
        new_files = [file for file in files if file not in self.transcript_files]
        self.transcript_files.extend(new_files)
        self.update_status_text(
            f"Selected {len(new_files)} new file(s). Total files: {len(self.transcript_files)}"
        )

    def update_status_text(self, message, tag=None):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n", tag)
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)

    def process_transcripts(self):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)

        success_count = 0
        failure_count = 0

        for file in self.transcript_files:
            try:
                transcript = self.read_docx(file)
                if self.is_processing_successful(transcript):
                    processed_data = self.process_transcript(transcript)
                    output_file = os.path.splitext(file)[0] + ".xlsx"
                    processed_data.to_excel(output_file, index=False)
                    self.update_status_text(
                        f"Successfully processed: {os.path.basename(file)}", "success"
                    )
                    success_count += 1
                else:
                    self.update_status_text(
                        f"Failed to process: {os.path.basename(file)}", "failure"
                    )
                    failure_count += 1
            except Exception as e:
                self.update_status_text(
                    f"Error processing {os.path.basename(file)}: {str(e)}", "failure"
                )
                failure_count += 1

        summary = f"\nProcessing complete.\nSuccessful: {success_count}\nFailed: {failure_count}"
        self.update_status_text(summary)

    def read_docx(self, file):
        doc = Document(file)
        full_text = [para.text for para in doc.paragraphs]
        return "\n".join(full_text)

    def is_processing_successful(self, transcript):
        try:
            processed_data = self.process_transcript(transcript)
            return (
                processed_data is not None
                and isinstance(processed_data, pd.DataFrame)
                and not processed_data.empty
            )
        except:
            return False

    def process_transcript(self, transcript):
        data = {
            "Speaker": [],
            "Teacher (T) or Child (C)": [],
            "Utterance/Idea Units": [],
        }
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

    def apply_template_with_direct_save(self, raw_data, template_file_path, save_path):
        try:
            with xw.App(visible=False) as app:
                template_wb = app.books.open(template_file_path)
                template_ws = template_wb.sheets[0]

                raw_data_list = raw_data.values.tolist()
                template_ws.range("A2").value = raw_data_list

                template_wb.save(save_path)
                template_wb.close()

            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error during applying template: {str(e)}")
            return False


def main():
    root = tk.Tk()
    app = ExcelTemplateApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
