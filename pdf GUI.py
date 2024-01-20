import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox

class PDFManipulator:
    def __init__(self, master):
        self.master = master
        master.title("PDF Manipulator")

        self.label = tk.Label(master, text="Kies een functionaliteit:")
        self.label.pack()

        self.split_button = self.create_button(master, "PDF splitsen", self.choose_split)
        self.merge_button = self.create_button(master, "PDF samenvoegen", self.choose_merge)

        self.error_message = None

    def create_button(self, master, text, command):
        button = tk.Button(master, text=text, command=command)
        button.pack()
        return button

    def choose_split(self):
        self.master.withdraw()
        self.split_window = self.create_toplevel("PDF Splitsen")

        self.label = tk.Label(self.split_window, text="Selecteer een PDF-bestand:")
        self.label.pack()

        self.file_path = None

        self.select_button = self.create_button(self.split_window, "Selecteer PDF", self.select_pdf)
        self.page_groups_label = tk.Label(self.split_window, text="Pagina's om te splitsen (bijv. 1, 3, 5-7, 10-end):")
        self.page_groups_label.pack()
        self.page_groups_entry = tk.Entry(self.split_window)
        self.page_groups_entry.pack()

        self.split_instruction_label = tk.Label(self.split_window, text="Voer paginanummers in (bijv. 1, 3, 5-7, 10-end)")
        self.split_instruction_label.pack()

        self.split_output_label = tk.Label(self.split_window, text="Selecteer een uitvoermap:")
        self.split_output_label.pack()

        self.split_output_entry = tk.Entry(self.split_window)
        self.split_output_entry.pack()

        self.select_output_button = self.create_button(self.split_window, "Selecteer map", lambda: self.select_output_path(self.split_output_entry))
        self.split_button = self.create_button(self.split_window, "PDF splitsen", self.split_pdf)
        self.back_button = self.create_button(self.split_window, "Terug naar keuzescherm", self.back_to_main)

    def choose_merge(self):
        self.master.withdraw()
        self.merge_window = self.create_toplevel("PDF Samenvoegen")

        self.main_label = tk.Label(self.merge_window, text="Selecteer het hoofddocument:")
        self.main_label.pack()

        self.main_file_path = None

        self.main_select_button = self.create_button(self.merge_window, "Selecteer hoofddocument", self.select_main_pdf)
        self.extra_label = tk.Label(self.merge_window, text="Selecteer extra documenten:")
        self.extra_label.pack()

        self.extra_file_paths = []

        self.extra_select_button = self.create_button(self.merge_window, "Selecteer extra documenten", self.select_extra_pdfs)
        self.merge_options_label = tk.Label(self.merge_window, text="Selecteer samenvoegopties:")
        self.merge_options_label.pack()

        self.merge_all_var = tk.IntVar()
        self.merge_all_checkbox = tk.Checkbutton(self.merge_window, text="Hele documenten samenvoegen", variable=self.merge_all_var, command=self.toggle_merge_options)
        self.merge_all_checkbox.pack()

        self.merge_ranges_label = tk.Label(self.merge_window, text="Paginaranges per extra document:")
        self.merge_ranges_label.pack()

        self.merge_ranges_entries = []

        self.select_ranges_button = self.create_button(self.merge_window, "Selecteer paginaranges", self.select_ranges)
        self.merge_output_label = tk.Label(self.merge_window, text="Selecteer een uitvoermap:")
        self.merge_output_label.pack()

        self.merge_output_entry = tk.Entry(self.merge_window)
        self.merge_output_entry.pack()

        self.select_output_button = self.create_button(self.merge_window, "Selecteer map", lambda: self.select_output_path(self.merge_output_entry))
        self.merge_button = self.create_button(self.merge_window, "PDF samenvoegen", self.merge_pdfs)
        self.back_button = self.create_button(self.merge_window, "Terug naar keuzescherm", self.back_to_main)

    def create_toplevel(self, title):
        toplevel = tk.Toplevel(self.master)
        toplevel.title(title)
        return toplevel

    def toggle_merge_options(self):
        for entry in self.merge_ranges_entries:
            entry.pack_forget() if self.merge_all_var.get() == 1 else entry.pack()

    def select_output_path(self, entry):
        output_path = filedialog.askdirectory(title="Selecteer uitvoermap")
        if output_path:
            entry.delete(0, tk.END)
            entry.insert(0, output_path)

    def select_pdfs(self, main=False):
        num_files = 1 if main else 2
        file_paths = filedialog.askopenfilenames(title=f"Selecteer {'hoofd' if main else 'extra'} PDF-bestand(en)", filetypes=[("PDF-bestanden", "*.pdf")])
        if file_paths and len(file_paths) >= num_files:
            if main:
                self.main_file_path = file_paths[0]
                self.main_label.config(text=f"Geselecteerd hoofddocument: {self.main_file_path}")
            else:
                self.extra_file_paths = file_paths
                self.extra_label.config(text=f"Geselecteerde extra documenten: {', '.join(self.extra_file_paths)}")
        else:
            self.show_error(f"Selecteer minstens {num_files} {'hoofd' if main else 'extra'} PDF-bestand(en).")

    def select_main_pdf(self):
        self.select_pdfs(main=True)

    def select_extra_pdfs(self):
        num_files = 2

        if self.merge_all_var.get() == 1:
            self.select_pdfs()
        else:
            if not self.main_file_path:
                self.show_error("Selecteer eerst het hoofddocument.")
                return

            self.extra_file_paths = []

            for i in range(num_files):
                file_path = filedialog.askopenfilename(title=f"Selecteer extra PDF-bestand {i + 1}", filetypes=[("PDF-bestanden", "*.pdf")])
                if file_path:
                    self.extra_file_paths.append(file_path)

            self.extra_label.config(text=f"Geselecteerde extra documenten: {', '.join(self.extra_file_paths)}")

    def select_ranges(self):
        if not self.extra_file_paths:
            self.show_error("Selecteer eerst extra documenten.")
            return

        self.merge_ranges_entries = []

        for i, file_path in enumerate(self.extra_file_paths):
            range_label = tk.Label(self.merge_window, text=f"Range voor Extra Document {i + 1}:")
            range_label.pack()

            range_entry = tk.Entry(self.merge_window)
            range_entry.pack()

            self.merge_ranges_entries.append(range_entry)

    def select_pdf(self):
        self.file_path = filedialog.askopenfilename(title="Selecteer PDF", filetypes=[("PDF-bestanden", "*.pdf")])
        if self.file_path:
            self.label.config(text=f"Geselecteerd PDF-bestand: {self.file_path}")

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

                output_file = f"{output_path}/group_{start}_{end}.pdf"
                with open(output_file, "wb") as output:
                    pdf_writer.write(output)

            self.show_info("PDF is succesvol gesplitst!")

        except IndexError:
            self.show_error("Ongeldige paginabereiken. Zorg ervoor dat de opgegeven pagina's binnen het bereik vallen.")

    def merge_pdfs(self):
        output_path = self.merge_output_entry.get()

        if not self.main_file_path:
            self.show_error("Selecteer eerst het hoofddocument.")
            return

        if not self.extra_file_paths:
            self.show_error("Selecteer eerst extra documenten.")
            return

        try:
            pdf_writer = PyPDF2.PdfWriter()

            # Hoofddocument
            pdf_reader_main = PyPDF2.PdfReader(self.main_file_path)
            for page_num in range(len(pdf_reader_main.pages)):
                pdf_writer.add_page(pdf_reader_main.pages[page_num])

            # Extra documenten
            for i, file_path in enumerate(self.extra_file_paths):
                pdf_reader_extra = PyPDF2.PdfReader(file_path)

                if self.merge_all_var.get() == 1:  # Alle pagina's samenvoegen
                    for page_num in range(len(pdf_reader_extra.pages)):
                        pdf_writer.add_page(pdf_reader_extra.pages[page_num])
                else:  # Geselecteerde paginaranges samenvoegen
                    page_ranges = self.merge_ranges_entries[i].get().split(',')
                    for page_range in page_ranges:
                        if '-' in page_range:
                            start, end = map(int, page_range.split('-'))
                        elif page_range.lower() == 'end':
                            start, end = 1, len(pdf_reader_extra.pages)
                        else:
                            start = end = int(page_range)

                        for page_num in range(start - 1, end):
                            pdf_writer.add_page(pdf_reader_extra.pages[page_num])

            output_file = filedialog.asksaveasfilename(
                title="Opslaan als",
                defaultextension=".pdf",
                initialdir=output_path,
                filetypes=[("PDF-bestanden", "*.pdf")]
            )

            if output_file:
                with open(output_file, "wb") as output:
                    pdf_writer.write(output)

                self.show_info("Geselecteerde pagina's zijn succesvol samengevoegd!")

        except IndexError:
            self.show_error("Ongeldige paginabereiken. Zorg ervoor dat de opgegeven pagina's binnen het bereik vallen.")

    def back_to_main(self):
        if hasattr(self, 'split_window'):
            self.split_window.destroy()
        elif hasattr(self, 'merge_window'):
            self.merge_window.destroy()

        self.master.deiconify()

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
