import PyPDF2
from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os


def show_error(message):
    messagebox.showerror("Fout", message)


class PDFManipulator:

    def __init__(self, window):
        self.window = window
        window.title("PDF Manipulator")

        self.label = tk.Label(master=window, text="Kies een functionaliteit:")
        self.label.grid(row=0, column=0)

        self.split_button = tk.Button(master=window, text="PDF splitsen", command=self.choose_split)
        self.split_button.grid(row=1, column=0, pady=10)

        self.decrypt_button = tk.Button(master=window, text="Beveiliging verwijderen", command=self.choose_decrypt)
        self.decrypt_button.grid(row=2, column=0, pady=10)

        self.merge_button = tk.Button(master=window, text="PDF samenvoegen", command=self.merge_pdf)
        self.merge_button.grid(row=3, column=0, pady=10)

        self.pdf_to_docx_button = tk.Button(master=window, text="PDF naar DOCX Converteren",
                                            command=self.choose_pdf_to_docx)
        self.pdf_to_docx_button.grid(row=4, column=0, pady=10)

        self.files = []
        self.page_ranges = {}

        self.error_message = None

    def choose_split(self):
        self.window.withdraw()
        self.split_window = tk.Tk()
        self.split_window.title("PDF Splitsen")

        self.label = tk.Label(self.split_window, text="PDF-bestand:")
        self.label.grid(row=0, column=0, sticky='e')

        self.file_path = None

        self.select_doc_entry = tk.Entry(self.split_window)
        self.select_doc_entry.grid(row=0, column=1)
        self.select_button = tk.Button(master=self.split_window, text="Selecteer PDF",
                                       command=lambda: self.select_pdf(self.select_doc_entry))
        self.select_button.grid(row=0, column=2)

        self.page_groups_label = tk.Label(self.split_window, text="Te splitsen pagina's (bijv. 1-1, 3-3, 5-7, 10-end):")
        self.page_groups_label.grid(row=2, column=0, sticky='e')
        self.page_groups_entry = tk.Entry(self.split_window)
        self.page_groups_entry.grid(row=2, column=1)

        self.split_output_label = tk.Label(self.split_window, text="Uitvoermap:")
        self.split_output_label.grid(row=1, column=0, sticky='e')
        self.split_output_entry = tk.Entry(self.split_window)
        self.split_output_entry.grid(row=1, column=1)
        self.select_output_button = tk.Button(master=self.split_window, text="Selecteer map",
                                              command=lambda: self.select_output_path(self.split_output_entry))
        self.select_output_button.grid(row=1, column=2)

        self.split_button = tk.Button(master=self.split_window, text="PDF splitsen", command=self.split_pdf)
        self.split_button.grid(row=3, column=2, padx=5, pady=2)

        self.back_button = tk.Button(master=self.split_window, text="Terug naar keuzescherm",
                                     command=self.back_to_main)
        self.back_button.grid(row=4, column=2, padx=5, pady=0)


    def choose_decrypt(self):
        self.window.withdraw()
        self.restrictions_window = tk.Toplevel(self.window)
        self.restrictions_window.title("Beperkingen Verwijderen")

        # Label en button voor het selecteren van een PDF-bestand
        self.label = tk.Label(self.restrictions_window, text="Selecteer een PDF-bestand:")
        self.label.grid(row=0, column=0, sticky='e')

        self.file_path = None  # Variable voor opslaan van geselecteerde bestandspad

        self.restrictions_pdf_entry = tk.Entry(self.restrictions_window)
        self.restrictions_pdf_entry.grid(row=0, column=1)

        self.select_button = tk.Button(master=self.restrictions_window, text="Selecteer PDF",
                                       command=lambda: self.select_pdf(self.restrictions_pdf_entry))
        self.select_button.grid(row=0, column=2)

        # Label en entry voor het selecteren van een uitvoermap
        self.restrictions_output_label = tk.Label(self.restrictions_window, text="Selecteer een uitvoermap:")
        self.restrictions_output_label.grid(row=1, column=0, sticky='e')

        self.restrictions_output_entry = tk.Entry(self.restrictions_window)
        self.restrictions_output_entry.grid(row=1, column=1)

        self.select_output_button = tk.Button(master=self.restrictions_window, text="Selecteer map",
                                              command=lambda: self.select_output_path(self.restrictions_output_entry))
        self.select_output_button.grid(row=1, column=2)

        # Button voor het starten van het verwijderen van beperkingen
        self.restrictions_button = tk.Button(master=self.restrictions_window, text="Verwijder beperkingen",
                                             command=self.remove_restrictions)
        self.restrictions_button.grid(row=2, column=2, padx=5, pady=2)

        # Terug button
        self.back_button = tk.Button(master=self.restrictions_window, text="Terug naar keuzescherm",
                                     command=self.back_to_main)
        self.back_button.grid(row=3, column=2, padx=5, pady=0)

    def choose_pdf_to_docx(self):
        self.window.withdraw()
        self.pdf_to_docx_window = tk.Tk()
        self.pdf_to_docx_window.title("PDF naar DOCX Converteren")

        # Knop voor het starten van de conversie
        self.convert_button = tk.Button(master=self.pdf_to_docx_window, text="Converteer naar DOCX",
                                   command=self.pdf_to_docx)
        self.convert_button.grid(row=2, column=2, padx=5, pady=2)

        # Terug naar hoofdmenu knop
        self.back_button = tk.Button(master=self.pdf_to_docx_window, text="Terug naar keuzescherm",
                                command=self.back_to_main)
        self.back_button.grid(row=3, column=2, padx=5, pady=0)

    def pdf_to_docx(self):
        input_pdf = filedialog.askopenfilename(
            title="Selecteer een PDF-bestand",
            filetypes=[("PDF-bestanden", "*.pdf")])
        if not input_pdf:
            return

        output_docx = filedialog.asksaveasfilename(
            title="Opslaan als DOCX",
            defaultextension=".docx",
            filetypes=[("DOCX-bestanden", "*.docx")])
        if not output_docx:
            return

        try:
            cv = Converter(input_pdf)
            cv.convert(output_docx, start=0, end=None)
            cv.close()
            messagebox.showinfo("Succes", "Conversie voltooid!")
        except Exception as e:
            messagebox.showerror("Fout bij conversie", str(e))


    def merge_pdf(self):
        self.files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not self.files:
            return

        self.page_ranges = {}
        for file in self.files:
            with open(file, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                pages = simpledialog.askstring("Pagina's selecteren",
                                               f"Selecteer pagina's voor {os.path.basename(file)} (bv. 1-3, 5, 18-end):")
                self.page_ranges[file] = self.parse_page_ranges(pages, total_pages)

        output_filename = simpledialog.askstring("Output bestandsnaam", "Voer de bestandsnaam in:")
        if not output_filename:
            return

        output_folder = filedialog.askdirectory()
        if not output_folder:
            return

        output_path = os.path.join(output_folder, output_filename + '.pdf')
        merger = PyPDF2.PdfMerger()

        for file, pages in self.page_ranges.items():
            with open(file, 'rb') as f:
                for page in pages:
                    merger.append(fileobj=f, pages=(page, page + 1))

        with open(output_path, 'wb') as f:
            merger.write(f)

        messagebox.showinfo("Klaar", f"Bestanden samengevoegd in {output_path}")

    def parse_page_ranges(self, ranges_str, total_pages):
        ranges = []
        for part in ranges_str.split(','):
            if '-' in part:
                start_str, end_str = part.split('-')
                start = int(start_str) - 1 if start_str.strip() else 0
                end = int(end_str) - 1 if end_str.strip().lower() != 'end' else total_pages - 1
                ranges.extend(range(start - 1, end))
            else:
                page = int(part.strip()) - 1
                ranges.append(page)
        return ranges

    def select_output_path(self, entry):
        self.output_path = filedialog.askdirectory(title="Selecteer uitvoermap")
        if self.output_path:
            entry.delete(0, tk.END)
            entry.insert(0, self.output_path)

    def select_pdf(self, entry):
        file_path = filedialog.askopenfilename(title="Selecteer PDF", filetypes=[("PDF-bestanden", "*.pdf")])
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)
            self.file_path = file_path

    def split_pdf(self):
        page_groups = self.page_groups_entry.get().split(',')
        output_path = self.split_output_entry.get()

        pdf_reader = PyPDF2.PdfReader(self.file_path)

        try:
            for group in page_groups:
                if '-' in group:
                    start_str, end_str = map(str.strip, group.split('-'))
                    start = int(start_str) if start_str else 1
                    end = int(end_str) if end_str and end_str.lower() != 'end' else len(pdf_reader.pages)
                else:
                    start = end = int(group) if group.isdigit() else 1

                pdf_writer = PyPDF2.PdfWriter()
                for page_num in range(start - 1, end):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

                output_file = f"{output_path}/splitsing_pagina_{start}_{end}.pdf"
                with open(output_file, "wb") as output:
                    pdf_writer.write(output)

            self.select_doc_entry.delete(0, tk.END)
            self.page_groups_entry.delete(0, tk.END)
            self.split_output_entry.delete(0, tk.END)
            self.show_info("PDF is succesvol gesplitst!")

        except IndexError:
            show_error("Ongeldige paginabereiken. Zorg ervoor dat de opgegeven pagina's binnen het bereik vallen.")


    def remove_restrictions(self):
        if not self.file_path or not self.restrictions_output_entry.get():
            show_error("Selecteer eerst een PDF-bestand en een uitvoermap.")
            return

        try:
            pdf_reader = PyPDF2.PdfReader(self.file_path)
            pdf_writer = PyPDF2.PdfWriter()

            for page_num in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page_num])

            output_file = os.path.join(self.restrictions_output_entry.get(), os.path.basename(self.file_path))

            if output_file:
                with open(output_file, "wb") as output:
                    pdf_writer.write(output)
                self.restrictions_pdf_entry.delete(0, tk.END)
                self.restrictions_output_entry.delete(0, tk.END)
                self.show_info("Beperkingen succesvol verwijderd!")

        except Exception as e:
            show_error(f"Fout bij verwijderen beperkingen: {str(e)}")

    def back_to_main(self):
        try:
            if hasattr(self, 'split_window') and self.split_window.winfo_exists():
                self.split_window.destroy()
        except tk.TclError:
            pass

        try:
            if hasattr(self, 'restrictions_window') and self.restrictions_window.winfo_exists():
                self.restrictions_window.destroy()
        except tk.TclError:
            pass

        try:
            if hasattr(self, 'pdf_to_docx_window') and self.pdf_to_docx_window.winfo_exists():
                self.pdf_to_docx_window.destroy()
        except tk.TclError:
            pass

        self.window.deiconify()

    def show_info(self, message):
        self.show_message("Succes", message)

    def show_message(self, title, message, method=messagebox.showinfo):
        if self.error_message:
            self.error_message.destroy()

        method(title, message)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFManipulator(root)
    root.mainloop()
