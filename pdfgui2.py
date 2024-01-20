import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox


class PDFManipulator:
    def __init__(self, window):
        self.window = window
        window.title("PDF Manipulator")

        self.label = tk.Label(master=window, text="Kies een functionaliteit:")
        self.label.grid(row=0, column=0)

        self.split_button = tk.Button(master=window, text="PDF splitsen", command=self.choose_split)
        self.split_button.grid(row=1, column=0, padx=2, pady=2)

        self.merge_button = tk.Button(master=window, text="PDF samenvoegen", command=self.choose_merge)
        self.merge_button.grid(row=2, column=0, padx=2, pady=2)

        self.restrictions_button = tk.Button(master=window, text="Beperkingen Verwijderen",
                                             command=self.choose_restrictions)
        self.restrictions_button.grid(row=3, column=0, padx=2, pady=2)

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

        self.page_groups_label = tk.Label(self.split_window, text="Te splitsen pagina's (bijv. 1, 3, 5-7, 10-end):")
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

    def choose_merge(self):
        self.window.withdraw()
        self.merge_window = tk.Tk()
        self.merge_window.title("PDF Samenvoegen")

        self.label = tk.Label(self.merge_window, text="Selecteer PDF-bestanden om samen te voegen:")
        self.label.grid(row=0, column=0, columnspan=4, pady=5)

        self.files_to_merge = []
        self.pages_per_document = {}
        self.page_ranges_per_document = {}

        self.merge_listbox = tk.Listbox(self.merge_window, selectmode=tk.MULTIPLE)
        self.merge_listbox.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

        self.add_files_button = tk.Button(self.merge_window, text="Bestanden toevoegen", command=self.add_merge_files)
        self.add_files_button.grid(row=2, column=0, padx=5, pady=5)

        self.remove_file_button = tk.Button(self.merge_window, text="Bestand verwijderen",
                                            command=self.remove_merge_file)
        self.remove_file_button.grid(row=2, column=1, padx=5, pady=5)

        self.page_info_label = tk.Label(self.merge_window, text="Aantal pagina's:")
        self.page_info_label.grid(row=3, column=0, columnspan=2, pady=5)

        self.page_range_label = tk.Label(self.merge_window, text="Gewenste pagina's (bijv. 1, 3, 5-end):")
        self.page_range_label.grid(row=4, column=0, pady=5)
        self.page_range_entry = tk.Entry(self.merge_window)
        self.page_range_entry.grid(row=4, column=1, pady=5)

        self.add_page_range_button = tk.Button(self.merge_window, text="Paginabereik toevoegen",
                                               command=self.add_page_range)
        self.add_page_range_button.grid(row=4, column=2, pady=5)

        self.output_label = tk.Label(self.merge_window, text="Uitvoerbestand:")
        self.output_label.grid(row=5, column=0, pady=5)
        self.output_entry = tk.Entry(self.merge_window)
        self.output_entry.grid(row=5, column=1, pady=5)
        self.select_output_button = tk.Button(self.merge_window, text="Selecteer bestand",
                                              command=lambda: self.select_output_path(self.output_entry))
        self.select_output_button.grid(row=5, column=2, pady=5)

        self.merge_button = tk.Button(self.merge_window, text="PDF samenvoegen", command=self.merge_pdfs)
        self.merge_button.grid(row=6, column=2, padx=5, pady=5)

        self.back_button = tk.Button(self.merge_window, text="Terug naar keuzescherm", command=self.back_to_main)
        self.back_button.grid(row=7, column=2, padx=5, pady=5)

# Naast Listbox steeds het totaal aantal pagina's per document weergeven
# Pagina range moet per document opgegeven kunnen worden op een voor de gebruiker duidelijk manier, het aantal mogelijk in te geven ranges moet overeenkomen met het aantal geselecteerde bestanden

    def choose_restrictions(self):
        self.window.withdraw()
        self.restrictions_window = tk.Tk()
        self.restrictions_window.title("Beperkingen Verwijderen")

        self.label = tk.Label(self.restrictions_window, text="Selecteer een PDF-bestand:")
        self.label.grid(row=0, column=0, sticky='e')

        self.file_path = None

        self.select_button = tk.Button(master=self.restrictions_window, text="Selecteer PDF",
                                       command=lambda: self.select_pdf(self.restrictions_pdf_entry))
        self.select_button.grid(row=0, column=2)

        self.restrictions_pdf_entry = tk.Entry(self.restrictions_window)
        self.restrictions_pdf_entry.grid(row=0, column=1)

        self.restrictions_output_label = tk.Label(self.restrictions_window, text="Selecteer een uitvoermap:")
        self.restrictions_output_label.grid(row=1, column=0, sticky='e')

        self.restrictions_output_entry = tk.Entry(self.restrictions_window)
        self.restrictions_output_entry.grid(row=1, column=1)

        self.select_output_button = tk.Button(master=self.restrictions_window, text="Selecteer map",
                                              command=lambda: self.select_output_path(self.restrictions_output_entry))
        self.select_output_button.grid(row=1, column=2)

        self.restrictions_button = tk.Button(master=self.restrictions_window, text="Verwijder beperkingen",
                                             command=self.remove_restrictions)
        self.restrictions_button.grid(row=2, column=2, padx=5, pady=2)

        self.back_button = tk.Button(master=self.restrictions_window, text="Terug naar keuzescherm",
                                     command=self.back_to_main)
        self.back_button.grid(row=3, column=2, padx=5, pady=0)

    def add_merge_files(self):
        files = filedialog.askopenfilenames(title="Selecteer PDF-bestanden", filetypes=[("PDF-bestanden", "*.pdf")])
        if files:
            self.files_to_merge.extend(files)
            for file in files:
                self.merge_listbox.insert(tk.END, file)
                self.update_pages_info(file)

    def remove_merge_file(self):
        selected_indices = self.merge_listbox.curselection()
        for idx in reversed(selected_indices):
            file_to_remove = self.files_to_merge.pop(idx)
            self.merge_listbox.delete(idx)
            del self.pages_per_document[file_to_remove]

    def update_pages_info(self, file_path):
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            self.pages_per_document[file_path] = len(pdf_reader.pages)

        total_pages = sum(self.pages_per_document.values())
        self.page_info_label.config(text=f"Aantal pagina's: {total_pages}")

    def add_page_range(self):
        selected_files = self.merge_listbox.curselection()
        if not selected_files:
            self.show_error("Selecteer minstens één bestand om een paginabereik toe te voegen.")
            return

        page_range = self.page_range_entry.get()
        if not page_range:
            self.show_error("Voer een paginabereik in.")
            return

        for file_index in selected_files:
            file_path = self.files_to_merge[file_index]
            if file_path not in self.pages_per_document:
                self.update_pages_info(file_path)

            pages = self.parse_page_range(page_range, self.pages_per_document[file_path])
            if 'end' in page_range.lower():
                self.show_info("Let op: 'end' wordt geïnterpreteerd als het einde van het document.")

            if file_path not in self.page_ranges_per_document:
                self.page_ranges_per_document[file_path] = []

            self.page_ranges_per_document[file_path].append(pages)

        self.page_range_entry.delete(0, tk.END)

    def parse_page_range(self, page_range, total_pages):
        pages = []
        groups = page_range.split(',')
        for group in groups:
            if '-' in group:
                start_str, end_str = map(str.strip, group.split('-'))
                start = int(start_str) if start_str else 1
                end = int(end_str) if end_str and end_str.lower() != 'end' else total_pages
            else:
                start = end = int(group) if group.isdigit() else 1

            pages.extend(range(start, end + 1))
        return pages

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
            self.show_error("Ongeldige paginabereiken. Zorg ervoor dat de opgegeven pagina's binnen het bereik vallen.")

    def merge_pdfs(self):
        page_range = self.page_range_entry.get()
        output_path = self.output_entry.get()

        try:
            pdf_merger = PyPDF2.PdfMerger()

            for file_path in self.files_to_merge:
                pdf_merger.append(file_path, pages=self.page_ranges_per_document.get(file_path, None))

            if page_range:
                pages_to_merge = self.parse_page_range(page_range)
                pdf_writer = PyPDF2.PdfWriter()

                for page_num in pages_to_merge:
                    pdf_writer.add_page(pdf_merger.pages[page_num - 1])

                with open(f"{output_path}/samenvoegen_pagina_{page_range}.pdf", "wb") as output:
                    pdf_writer.write(output)
            else:
                with open(f"{output_path}/samenvoegen_volledig.pdf", "wb") as output:
                    pdf_merger.write(output)

            self.page_range_entry.delete(0, tk.END)
            self.output_entry.delete(0, tk.END)
            self.show_info("PDF's zijn succesvol samengevoegd!")

        except Exception as e:
            self.show_error(f"Fout bij samenvoegen PDF's: {str(e)}")

    def remove_restrictions(self):
        if not self.file_path:
            self.show_error("Selecteer eerst een PDF-bestand.")
            return

        output_path = self.restrictions_output_entry.get()

        try:
            pdf_reader = PyPDF2.PdfReader(self.file_path)
            pdf_writer = PyPDF2.PdfWriter()

            for page_num in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page_num])

            output_file = filedialog.asksaveasfilename(
                title="Opslaan als",
                defaultextension=".pdf",
                initialdir=output_path,
                filetypes=[("PDF-bestanden", "*.pdf")]
            )

            if output_file:
                with open(output_file, "wb") as output:
                    pdf_writer.write(output)
                self.restrictions_pdf_entry.delete(0, tk.END)
                self.restrictions_output_entry.delete(0, tk.END)
                self.show_info("Beperkingen succesvol verwijderd!")

        except Exception as e:
            self.show_error(f"Fout bij verwijderen beperkingen: {str(e)}")

    def back_to_main(self):
        if hasattr(self, 'split_window'):
            self.split_window.destroy()
        if hasattr(self, 'merge_window'):
            self.merge_window.destroy()
        if hasattr(self, 'restrictions_window'):
            self.restrictions_window.destroy()

        self.window.deiconify()

    def show_info(self, message):
        self.show_message("Succes", message)

    def show_error(self, message):
        self.show_message("Fout", message, messagebox.showerror)

    def show_message(self, title, message, method=messagebox.showinfo):
        if self.error_message:
            self.error_message.destroy()

        method(title, message)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFManipulator(root)
    root.mainloop()

# Invoegen/vervangen pagina
# Foto's uit pdf extracten
# Tekst uit pdf trekken naar word