#!/usr/bin/env python3
"""PDF to DOCX Converter — preserves layout, fonts, tables, and images."""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
from pathlib import Path


def convert(pdf_path: str, out_path: str, progress_callback=None):
    from pdf2docx import Converter
    cv = Converter(pdf_path)
    cv.convert(out_path, start=0, end=None)
    cv.close()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF → DOCX Converter")
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")
        self._build_ui()
        self._center()

    def _build_ui(self):
        pad = dict(padx=20, pady=10)

        # Drop zone / file picker
        self.drop_frame = tk.Frame(
            self, bg="#e8eaf6", relief="flat", bd=2,
            highlightbackground="#9fa8da", highlightthickness=2,
            cursor="hand2"
        )
        self.drop_frame.pack(fill="x", padx=20, pady=(20, 5))
        self.drop_frame.bind("<Button-1>", lambda e: self._pick_file())

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Click to choose a PDF file",
            font=("Helvetica", 13), bg="#e8eaf6", fg="#3949ab", pady=28
        )
        self.drop_label.pack()
        self.drop_label.bind("<Button-1>", lambda e: self._pick_file())

        # Selected file display
        self.file_var = tk.StringVar(value="No file selected")
        self.file_label = tk.Label(
            self, textvariable=self.file_var,
            font=("Helvetica", 10), bg="#f5f5f5", fg="#555", wraplength=400
        )
        self.file_label.pack(**pad)

        # Convert button
        self.btn = tk.Button(
            self, text="Convert to DOCX", command=self._run,
            font=("Helvetica", 13, "bold"), bg="#3949ab", fg="white",
            activebackground="#283593", activeforeground="white",
            relief="flat", padx=24, pady=10, cursor="hand2"
        )
        self.btn.pack(pady=(0, 10))

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=400)
        self.progress.pack(padx=20, pady=(0, 5))

        # Status
        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(
            self, textvariable=self.status_var,
            font=("Helvetica", 10), bg="#f5f5f5", fg="#555"
        )
        self.status_label.pack(pady=(0, 20))

        self.pdf_path = None

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_path = path
            self.file_var.set(f"Selected: {Path(path).name}")
            self.drop_label.config(text="✓ PDF selected — click Convert")

    def _run(self):
        if not self.pdf_path:
            messagebox.showwarning("No file", "Please select a PDF first.")
            return

        default_name = Path(self.pdf_path).stem + ".docx"
        out_path = filedialog.asksaveasfilename(
            title="Save DOCX as",
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")]
        )
        if not out_path:
            return

        self.btn.config(state="disabled")
        self.status_var.set("Converting…")
        self.progress.start(12)

        def worker():
            try:
                convert(self.pdf_path, out_path)
                self.after(0, self._done, out_path, None)
            except Exception as exc:
                self.after(0, self._done, out_path, exc)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, out_path, error):
        self.progress.stop()
        self.btn.config(state="normal")
        if error:
            self.status_var.set(f"Error: {error}")
            messagebox.showerror("Conversion failed", str(error))
        else:
            self.status_var.set(f"Saved: {Path(out_path).name}")
            if messagebox.askyesno("Done!", f"Saved to:\n{out_path}\n\nOpen the file now?"):
                os.system(f'open "{out_path}"')

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    # CLI mode: python convert.py input.pdf [output.docx]
    if len(sys.argv) >= 2:
        pdf = sys.argv[1]
        out = sys.argv[2] if len(sys.argv) >= 3 else str(Path(pdf).with_suffix(".docx"))
        print(f"Converting {pdf} → {out}")
        convert(pdf, out)
        print("Done.")
    else:
        App().mainloop()
