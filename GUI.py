import tkinter as tk
from tkinter import filedialog, messagebox
from logic import setup_driver, get_links, get_financial_data, download_pdfs, convert_pdf_to_excel

class FinancialDataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Data Scraper")

        self.scrape_var = tk.BooleanVar()
        self.download_var = tk.BooleanVar()
        self.convert_var = tk.BooleanVar()

        self.create_widgets()

    def create_widgets(self):
        tk.Checkbutton(self.root, text="Scrape Financial Data", variable=self.scrape_var).pack()
        tk.Checkbutton(self.root, text="Download PDFs", variable=self.download_var).pack()
        tk.Checkbutton(self.root, text="Convert PDFs to Excel", variable=self.convert_var).pack()

        self.start_button = tk.Button(self.root, text="Start", command=self.start_process)
        self.start_button.pack()

        self.output_label = tk.Label(self.root, text="")
        self.output_label.pack()

    def start_process(self):
        driver = setup_driver()
        names_links_dict = get_links(driver)

        selected_companies = self.get_selected_companies(names_links_dict)

        if self.scrape_var.get():
            output_file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if output_file:
                get_financial_data(driver, names_links_dict, selected_companies, output_file)

        if self.download_var.get():
            download_pdfs(driver, names_links_dict, selected_companies)

        if self.convert_var.get():
            pdf_dir = filedialog.askdirectory()
            if pdf_dir:
                pdf_files = self.get_pdf_files(pdf_dir, selected_companies)
                convert_pdf_to_excel(pdf_files, pdf_dir)

        messagebox.showinfo("Process Completed", "The selected processes have been completed.")

    def get_selected_companies(self, names_links_dict):
        selected_companies = []
        select_window = tk.Toplevel(self.root)
        select_window.title("Select Companies")

        for company in names_links_dict.keys():
            var = tk.BooleanVar()
            tk.Checkbutton(select_window, text=company, variable=var).pack()
            selected_companies.append((company, var))

        def on_select():
            self.selected_companies = [company for company, var in selected_companies if var.get()]
            select_window.destroy()

        tk.Button(select_window, text="Select", command=on_select).pack()
        self.root.wait_window(select_window)

        return self.selected_companies

    def get_pdf_files(self, pdf_dir, selected_companies):
        pdf_files = []
        for company in selected_companies:
            company_dir = os.path.join(pdf_dir, company)
            for year in os.listdir(company_dir):
                if year.endswith(".pdf"):
                    pdf_files.append((os.path.join(company_dir, year), company, year.split(".")[0]))
        return pdf_files

if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialDataApp(root)
    root.mainloop()

