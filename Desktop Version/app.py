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
        master.geometry("600x400")

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.stage1_frame = ttk.Frame(self.notebook)
        self.stage2_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.stage1_frame, text="Stage 1: Transcript to Raw Excel")
        self.notebook.add(self.stage2_frame, text="Stage 2: Apply Template")

        self.setup_stage1()
        self.setup_stage2()

    def setup_stage1(self):
        ttk.Label(self.stage1_frame, text="Upload Transcript Files (.docx)").pack(
            pady=10
        )
        self.transcript_files = []
        self.transcript_listbox = tk.Listbox(self.stage1_frame, width=70, height=10)
        self.transcript_listbox.pack(pady=5)

        ttk.Button(
            self.stage1_frame, text="Select Files", command=self.load_transcripts
        ).pack(pady=5)
        ttk.Button(
            self.stage1_frame,
            text="Process Transcripts",
            command=self.process_transcripts,
        ).pack(pady=5)

    def setup_stage2(self):
        self.stage2_frame.columnconfigure(1, weight=1)
        self.stage2_frame.rowconfigure(5, weight=1)

        ttk.Label(self.stage2_frame, text="Template File:").grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )
        self.template_entry = ttk.Entry(self.stage2_frame, width=50)
        self.template_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(
            self.stage2_frame,
            text="Select",
            command=lambda: self.load_file(self.template_entry),
        ).grid(row=0, column=2, sticky="w", padx=5, pady=5)

        ttk.Label(self.stage2_frame, text="Raw Excel File:").grid(
            row=1, column=0, sticky="w", padx=5, pady=5
        )
        self.raw_entry = ttk.Entry(self.stage2_frame, width=50)
        self.raw_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(
            self.stage2_frame,
            text="Select",
            command=lambda: self.load_file(self.raw_entry, self.preview_raw_file),
        ).grid(row=1, column=2, sticky="w", padx=5, pady=5)

        ttk.Button(
            self.stage2_frame, text="Apply Template", command=self.apply_template
        ).grid(row=2, column=1, sticky="ew", padx=5, pady=10)

        ttk.Label(self.stage2_frame, text="Raw Excel File Preview:").grid(
            row=3, column=0, columnspan=3, sticky="w", padx=5, pady=5
        )
        self.preview_text = tk.Text(self.stage2_frame, wrap=tk.NONE, height=10)
        self.preview_text.grid(
            row=4, column=0, columnspan=3, sticky="nsew", padx=5, pady=5
        )

        y_scrollbar = ttk.Scrollbar(
            self.stage2_frame, orient="vertical", command=self.preview_text.yview
        )
        y_scrollbar.grid(row=4, column=3, sticky="ns")
        x_scrollbar = ttk.Scrollbar(
            self.stage2_frame, orient="horizontal", command=self.preview_text.xview
        )
        x_scrollbar.grid(row=5, column=0, columnspan=3, sticky="ew")

        self.preview_text.configure(
            yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set
        )

    def load_file(self, entry, callback=None):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            entry.delete(0, tk.END)
            entry.insert(0, file)
            if callback:
                callback(file)

    def preview_raw_file(self, file_path):
        try:
            df = pd.read_excel(file_path)
            preview = df.head().to_string()
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, preview)
        except Exception as e:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, f"Error previewing file: {str(e)}")

    def load_transcripts(self):
        files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")])
        new_files = [file for file in files if file not in self.transcript_files]
        self.transcript_files.extend(new_files)
        self.update_transcript_listbox()

    def update_transcript_listbox(self):
        self.transcript_listbox.delete(0, tk.END)
        for file in self.transcript_files:
            self.transcript_listbox.insert(tk.END, os.path.basename(file))

    def process_transcripts(self):
        for file in self.transcript_files:
            try:
                transcript = self.read_docx(file)
                processed_data = self.process_transcript(transcript)
                output_file = os.path.splitext(file)[0] + "_processed.xlsx"
                processed_data.to_excel(output_file, index=False)
                messagebox.showinfo("Success", f"Processed and saved: {output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process {file}: {str(e)}")

    def read_docx(self, file):
        doc = Document(file)
        full_text = [para.text for para in doc.paragraphs]
        return "\n".join(full_text)

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

    def apply_template(self):
        template_file = self.template_entry.get()
        raw_file = self.raw_entry.get()

        if not template_file or not raw_file:
            messagebox.showerror("Error", "Please select both template and raw files.")
            return

        try:
            raw_data = pd.read_excel(raw_file)
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")

            if save_path:
                success = self.apply_template_with_direct_save(
                    raw_data, template_file, save_path
                )
                if success:
                    messagebox.showinfo(
                        "Success", f"Template applied and saved to: {save_path}"
                    )
                else:
                    messagebox.showerror("Error", "Failed to apply template.")
            else:
                messagebox.showinfo("Info", "Save operation cancelled.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTemplateApp(root)
    root.mainloop()
