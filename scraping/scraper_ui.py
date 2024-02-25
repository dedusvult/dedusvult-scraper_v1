import subprocess
import sys
from package_installer import PackageInstaller

package_installer = PackageInstaller(["selenium", "webdriver_manager", "requests", "pandas", "openpyxl", "bs4"])
package_installer.install_packages()

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import global_vars
import threading
import os
from scraper import WebScraper


class TextRedirector(object):
    def __init__(self, widget):        self.widget = widget

    def write(self, string):
        self.widget.config(state='normal')
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)
        self.widget.config(state='disabled')

    def _write(self, string):
        self.widget.config(state='normal')
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)
        self.widget.config(state='disabled')

    def flush(self):
        pass


class ScraperUI:
    def __init__(self):
        self.output_text = None
        self.selected_tax_year = None
        self.confirm_button = None
        self.selected_zip_code = None
        self.selected_file_path = None
        self.selected_process_group = None
        self.root = tk.Tk()
        self.root.title("Scraper UI")
        self.root.geometry("350x200")  # Adjust the size as needed

        self.file_path_var = tk.StringVar()
        self.zip_code_var = tk.StringVar()
        self.process_group_var = tk.StringVar()
        self.tax_year_var = tk.StringVar(value=str(global_vars.tax_year))  # Directly use the variable

        self.setup_ui()

    def setup_ui(self):
        # Set a font
        default_font = ('Helvetica', 10)

        # Main window's starting position (centered)
        window_width = 400
        window_height = 250
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        self.root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

        # Style configuration
        style = ttk.Style()
        style.configure('TLabel', font=default_font)
        style.configure('TButton', font=default_font)
        style.configure('TEntry', font=default_font)
        style.configure('TMenubutton', font=default_font)

        self.file_selection_frame_setup(default_font)
        self.zip_code_frame_setup()
        self.process_group_frame_setup()
        self.tax_year_frame_setup()
        self.confirm_button_frame_setup()

    def confirm_button_frame_setup(self):
        # Confirm button
        confirm_frame = ttk.Frame(self.root, padding="10")
        confirm_frame.pack(fill='x')
        self.confirm_button = ttk.Button(confirm_frame, text="Confirm", command=self.confirm_selection,
                                         state='disabled')
        self.confirm_button.pack(pady=10, padx=10)
        # Update the confirm button state based on file selection
        self.file_path_var.trace("w", self.update_confirm_button_state)

    def tax_year_frame_setup(self):
        # Tax year selection frame
        tax_year_frame = ttk.Frame(self.root, padding="10")
        tax_year_frame.pack(fill='x', expand=True)
        # Spacers for centering the tax year dropdown
        left_spacer_year = ttk.Frame(tax_year_frame)
        left_spacer_year.pack(side='left', expand=True)
        ttk.Label(tax_year_frame, text="Tax Year:").pack(side='left')
        tax_year_options = [str(year) for year in range(2015, 2031)]
        tax_year_dropdown = ttk.OptionMenu(tax_year_frame, self.tax_year_var, str(global_vars.tax_year),
                                           *tax_year_options)
        tax_year_dropdown.pack(side='left')
        right_spacer_year = ttk.Frame(tax_year_frame)
        right_spacer_year.pack(side='left', expand=True)

    def zip_code_frame_setup(self):
        # Zip code selection frame
        zip_frame = ttk.Frame(self.root, padding="10")
        zip_frame.pack(fill='x', expand=True)
        # Spacers for centering the zip code dropdown
        left_spacer_zip = ttk.Frame(zip_frame)
        left_spacer_zip.pack(side='left', expand=True)
        ttk.Label(zip_frame, text="Zip Code:").pack(side='left')
        zip_code_options = ["Home Zip", "Work Zip"]
        zip_code_dropdown = ttk.OptionMenu(zip_frame, self.zip_code_var, zip_code_options[0], *zip_code_options)
        zip_code_dropdown.pack(side='left')
        right_spacer_zip = ttk.Frame(zip_frame)
        right_spacer_zip.pack(side='left', expand=True)

    def process_group_frame_setup(self):
        # Process group selection frame
        process_group_frame = ttk.Frame(self.root, padding="10")
        process_group_frame.pack(fill='x', expand=True)
        # Spacers for centering the process group dropdown
        left_spacer_zip = ttk.Frame(process_group_frame)
        left_spacer_zip.pack(side='left', expand=True)
        ttk.Label(process_group_frame, text="Process group:").pack(side='left')
        process_group_options = ["All employees", "FT and no status employees"]
        process_group_dropdown = ttk.OptionMenu(process_group_frame, self.process_group_var, process_group_options[0],
                                                *process_group_options)
        process_group_dropdown.pack(side='left')
        right_spacer_zip = ttk.Frame(process_group_frame)
        right_spacer_zip.pack(side='left', expand=True)

    def file_selection_frame_setup(self, default_font):
        # File selection frame
        file_frame = ttk.Frame(self.root, padding="10")
        file_frame.pack(fill='x', expand=True)
        ttk.Label(file_frame, text="Select File:", font=default_font).pack(side='left')
        ttk.Entry(file_frame, textvariable=self.file_path_var, font=default_font, width=25).pack(side='left', fill='x',
                                                                                                 expand=True)
        ttk.Button(file_frame, text="Browse", command=self.select_file, style='TButton').pack(side='right')

    def select_file(self):
        file_path = filedialog.askopenfilename()
        self.file_path_var.set(file_path)

    def update_confirm_button_state(self, *args):
        if self.file_path_var.get():
            self.confirm_button['state'] = 'normal'
        else:
            self.confirm_button['state'] = 'disabled'

    def confirm_selection(self):
        file_path = self.file_path_var.get()
        if file_path and os.path.exists(file_path):
            self.selected_file_path = file_path
            self.selected_zip_code = self.zip_code_var.get()
            self.selected_tax_year = self.tax_year_var.get()
            self.selected_process_group = self.process_group_var.get()

            # Create and configure the scraper
            scraper = WebScraper(self.selected_file_path, self.selected_zip_code, self.selected_tax_year)

            # Create output window and redirect output
            self.show_output_window()

            # Run the scraper in a separate thread
            scraper_thread = threading.Thread(target=scraper.run)
            scraper_thread.start()
        else:
            messagebox.showerror("Error", "File not selected or doesn't exist")
            self.file_path_var.set("")

    def show_output_window(self):
        output_window = tk.Toplevel(self.root)
        output_window.title("Output")
        output_window.geometry("800x600")

        output_text = scrolledtext.ScrolledText(output_window, state='disabled', height=12)
        output_text.pack(fill='both', expand=True)

        # Setup hyperlink manager

        # Redirect standard output to this new window with hyperlink capability
        sys.stdout = TextRedirector(output_text)
        self.output_text = output_text

    def run(self):
        self.root.mainloop()

        # Restore standard output
        sys.stdout = sys.__stdout__

        return self.selected_file_path, self.selected_zip_code, self.selected_tax_year


ScraperUI().run()
