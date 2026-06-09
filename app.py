import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None


WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_PAGE_BREAK = 7
WD_CHARACTER = 1
WD_STATISTIC_PAGES = 2
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST_PAGE = 2
WD_HEADER_FOOTER_EVEN_PAGES = 3


class WordPageExtractorError(Exception):
    pass


def require_windows_and_word_support():
    if os.name != "nt":
        raise WordPageExtractorError("This tool currently requires Windows.")
    if win32com is None or pythoncom is None:
        raise WordPageExtractorError(
            "Missing dependency: pywin32. Run: pip install -r requirements.txt"
        )


def normalize_docx_path(path: str) -> str:
    if not path:
        raise WordPageExtractorError("Please select a Word file.")
    if not os.path.exists(path):
        raise WordPageExtractorError("The selected Word file does not exist.")
    ext = os.path.splitext(path)[1].lower()
    if ext not in {".docx", ".docm", ".doc"}:
        raise WordPageExtractorError("Please select a Word document: .docx, .docm or .doc")
    return os.path.abspath(path)


def validate_page_range(start_page: int, end_page: int):
    if start_page < 1 or end_page < 1:
        raise WordPageExtractorError("Page numbers must be 1 or higher.")
    if end_page < start_page:
        raise WordPageExtractorError("End page must be equal to or higher than start page.")


def copy_page_setup(source_section, target_section):
    source_setup = source_section.PageSetup
    target_setup = target_section.PageSetup

    properties = [
        "TopMargin",
        "BottomMargin",
        "LeftMargin",
        "RightMargin",
        "HeaderDistance",
        "FooterDistance",
        "PageWidth",
        "PageHeight",
        "Orientation",
        "Gutter",
        "MirrorMargins",
        "DifferentFirstPageHeaderFooter",
        "OddAndEvenPagesHeaderFooter",
    ]

    for prop in properties:
        try:
            setattr(target_setup, prop, getattr(source_setup, prop))
        except Exception:
            # Some properties may not be available in all Word versions or document types.
            pass


def copy_headers_and_footers(source_section, target_section):
    header_footer_indices = [
        WD_HEADER_FOOTER_PRIMARY,
        WD_HEADER_FOOTER_FIRST_PAGE,
        WD_HEADER_FOOTER_EVEN_PAGES,
    ]

    for index in header_footer_indices:
        try:
            source_header = source_section.Headers(index)
            target_header = target_section.Headers(index)
            target_header.Range.FormattedText = source_header.Range.FormattedText
        except Exception:
            pass

        try:
            source_footer = source_section.Footers(index)
            target_footer = target_section.Footers(index)
            target_footer.Range.FormattedText = source_footer.Range.FormattedText
        except Exception:
            pass


def get_page_range(document, start_page: int, end_page: int):
    total_pages = document.ComputeStatistics(WD_STATISTIC_PAGES)
    if start_page > total_pages:
        raise WordPageExtractorError(
            f"Start page {start_page} is higher than the document page count ({total_pages})."
        )
    if end_page > total_pages:
        raise WordPageExtractorError(
            f"End page {end_page} is higher than the document page count ({total_pages})."
        )

    start_range = document.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=start_page)

    if end_page < total_pages:
        next_page_range = document.GoTo(
            What=WD_GOTO_PAGE,
            Which=WD_GOTO_ABSOLUTE,
            Count=end_page + 1,
        )
        page_range = document.Range(Start=start_range.Start, End=next_page_range.Start)
        # Avoid copying the page break/first character of the next page when Word exposes it in range.
        try:
            page_range.MoveEnd(WD_CHARACTER, -1)
        except Exception:
            pass
    else:
        page_range = document.Range(Start=start_range.Start, End=document.Content.End)

    return page_range


def extract_pages_to_docx(input_path: str, output_path: str, start_page: int, end_page: int):
    require_windows_and_word_support()
    input_path = normalize_docx_path(input_path)
    output_path = os.path.abspath(output_path)
    validate_page_range(start_page, end_page)

    pythoncom.CoInitialize()
    word = None
    source_doc = None
    target_doc = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        source_doc = word.Documents.Open(input_path, ReadOnly=True)
        source_doc.Repaginate()

        page_range = get_page_range(source_doc, start_page, end_page)
        source_section = page_range.Sections(1)

        target_doc = word.Documents.Add()
        target_section = target_doc.Sections(1)

        copy_page_setup(source_section, target_section)
        copy_headers_and_footers(source_section, target_section)

        target_range = target_doc.Range(0, 0)
        target_range.FormattedText = page_range.FormattedText

        # Remove a trailing page break if one was copied into the new document.
        try:
            text = target_doc.Range().Text
            if text.endswith("\f\r") or text.endswith("\f"):
                end_range = target_doc.Range(target_doc.Content.End - 2, target_doc.Content.End - 1)
                end_range.Delete()
        except Exception:
            pass

        target_doc.SaveAs2(output_path, FileFormat=16)  # wdFormatXMLDocument / DOCX
        return output_path

    finally:
        if target_doc is not None:
            try:
                target_doc.Close(False)
            except Exception:
                pass
        if source_doc is not None:
            try:
                source_doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


class WordPageExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word Page Extractor")
        self.geometry("620x330")
        self.resizable(False, False)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.start_page = tk.StringVar(value="1")
        self.end_page = tk.StringVar(value="1")
        self.status = tk.StringVar(value="Ready")

        self.build_ui()

    def build_ui(self):
        padding = {"padx": 12, "pady": 8}

        title = ttk.Label(
            self,
            text="Extract a Word page/range to a new editable DOCX",
            font=("Segoe UI", 13, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", **padding)

        ttk.Label(self, text="Source Word file:").grid(row=1, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.input_path, width=58).grid(row=1, column=1, sticky="we", **padding)
        ttk.Button(self, text="Browse", command=self.select_input).grid(row=1, column=2, **padding)

        ttk.Label(self, text="Start page:").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.start_page, width=12).grid(row=2, column=1, sticky="w", **padding)

        ttk.Label(self, text="End page:").grid(row=3, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.end_page, width=12).grid(row=3, column=1, sticky="w", **padding)

        ttk.Label(self, text="Output DOCX:").grid(row=4, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.output_path, width=58).grid(row=4, column=1, sticky="we", **padding)
        ttk.Button(self, text="Save as", command=self.select_output).grid(row=4, column=2, **padding)

        ttk.Button(self, text="Extract", command=self.run_extraction).grid(row=5, column=1, sticky="e", **padding)

        ttk.Separator(self).grid(row=6, column=0, columnspan=3, sticky="we", padx=12, pady=8)
        ttk.Label(self, textvariable=self.status, foreground="#333333").grid(
            row=7, column=0, columnspan=3, sticky="w", **padding
        )

        note = ttk.Label(
            self,
            text="Requires Windows + Microsoft Word installed. Complex tables/floating objects may need manual review.",
            wraplength=580,
        )
        note.grid(row=8, column=0, columnspan=3, sticky="w", **padding)

    def select_input(self):
        path = filedialog.askopenfilename(
            title="Select Word document",
            filetypes=[("Word documents", "*.docx *.docm *.doc"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            if not self.output_path.get():
                base, _ = os.path.splitext(path)
                self.output_path.set(f"{base}_extracted.docx")

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="Save extracted document as",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
        )
        if path:
            self.output_path.set(path)

    def run_extraction(self):
        try:
            input_path = self.input_path.get().strip()
            output_path = self.output_path.get().strip()
            if not output_path:
                raise WordPageExtractorError("Please choose an output file path.")

            start_page = int(self.start_page.get().strip())
            end_page = int(self.end_page.get().strip())

            self.status.set("Working...")
            self.update_idletasks()

            result = extract_pages_to_docx(input_path, output_path, start_page, end_page)
            self.status.set(f"Done: {result}")
            messagebox.showinfo("Success", f"Extracted document saved:\n{result}")

        except ValueError:
            self.status.set("Error: page numbers must be integers.")
            messagebox.showerror("Invalid page number", "Page numbers must be integers.")
        except Exception as exc:
            self.status.set(f"Error: {exc}")
            traceback.print_exc()
            messagebox.showerror("Extraction failed", str(exc))


if __name__ == "__main__":
    app = WordPageExtractorApp()
    app.mainloop()
