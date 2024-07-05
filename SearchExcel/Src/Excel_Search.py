import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.font import Font
import os
import fnmatch
import openpyxl
import tempfile
import subprocess
import configparser
import threading
import time
import xlrd
import csv
import chardet

DEBUG = False


class PathShortener:
    @staticmethod
    def shorten_path(path, max_length=None):
        path = os.path.normpath(path)
        if max_length is None:
            max_length = 100 if DEBUG else 50
        if len(path) <= max_length:
            return path

        path_parts = path.split(os.sep)
        shortened = path_parts[0] + os.sep + path_parts[1]
        for i in range(2, len(path_parts)):
            if len(shortened + os.sep + os.sep.join(path_parts[i:])) <= max_length:
                shortened = os.sep.join(path_parts[:i + 1])
            else:
                break
        while len(shortened) + 3 + len(os.sep + os.sep.join(path_parts[i:])) > max_length:
            i += 1
        return shortened + os.sep + "..." + os.sep + os.sep.join(path_parts[i:])

    @staticmethod
    def shorten_path_pixels(path, max_pixels=500, widget=None):
        def text_length_in_pixels(text):
            font = Font(font=widget.cget("font"))
            return font.measure(text)

        path = os.path.normpath(path)
        if text_length_in_pixels(path) <= max_pixels:
            return path

        path_parts = path.split(os.sep)
        shortened = path_parts[0] + os.sep
        i = 1
        while i < len(path_parts):
            next_part = shortened + os.sep + path_parts[i]
            if text_length_in_pixels(next_part + os.sep + '...') + text_length_in_pixels(os.sep.join(path_parts[-1:])) > max_pixels / 2:
                break
            shortened = next_part
            i += 1

        trailing = os.sep.join(path_parts[i:])
        while text_length_in_pixels(shortened + os.sep + '...' + os.sep + trailing) > max_pixels and i < len(path_parts):
            i += 1
            trailing = os.sep.join(path_parts[i:])
        return shortened + os.sep + '...' + os.sep + trailing


class ExcelSearcher:
    def __init__(self, base_folder, recursive=False):
        self.base_folder = base_folder
        self.recursive = recursive
        self.searching = False

    def search_excel_files(self, fname_match, progress_callback=None, include_csv=False):
        found_files = []

        def file_matches(file_name):
            return (
                fnmatch.fnmatch(file_name, f'*{fname_match}*.xlsx') or
                fnmatch.fnmatch(file_name, f'*{fname_match}*.xltx') or
                fnmatch.fnmatch(file_name, f'*{fname_match}*.xlsm') or
                fnmatch.fnmatch(file_name, f'*{fname_match}*.xls') or
                (include_csv and fnmatch.fnmatch(file_name, f'*{fname_match}*.csv'))
            )

        if self.recursive:
            for root, _, files in os.walk(self.base_folder):
                if progress_callback:
                    progress_callback(root)
                if not self.searching:
                    break
                for file_name in files:
                    if file_matches(file_name):
                        found_files.append(os.path.join(root, file_name))
        else:
            for subdir in os.listdir(self.base_folder):
                if not self.searching:
                    break
                subdir_path = os.path.join(self.base_folder, subdir)
                if os.path.isdir(subdir_path):
                    if progress_callback:
                        progress_callback(f"Scanning directories {subdir}")
                    for file_name in os.listdir(subdir_path):
                        if file_matches(file_name):
                            found_files.append(os.path.join(subdir_path, file_name))
        return found_files

    def search_excel(self, file_path, search_text):
        def detect_encoding(file_path):
            with open(file_path, 'rb') as f:
                result = chardet.detect(f.read())
            return result['encoding']

        found_rows = []

        if file_path.lower().endswith(('.xlsx', '.xltx', '.xlsm')):
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
            for row in sheet.iter_rows():
                if not self.searching:
                    break
                for cell in row:
                    if cell.value and search_text.lower() in str(cell.value).lower():
                        found_rows.append([cell.value for cell in row])
                        return found_rows
        elif file_path.lower().endswith('.xls'):
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            for row_idx in range(sheet.nrows):
                if not self.searching:
                    break
                row = sheet.row(row_idx)
                for cell in row:
                    cell_value = cell.value
                    if cell_value and search_text.lower() in str(cell_value).lower():
                        found_rows.append([cell.value for cell in row])
                        return found_rows
        elif file_path.lower().endswith('.csv'):
            encoding = detect_encoding(file_path)
            with open(file_path, 'r', newline='', encoding=encoding) as csvfile:
                csvreader = csv.reader(csvfile, delimiter=';')
                for row in csvreader:
                    if not self.searching:
                        break
                    for cell in row:
                        if search_text.lower() in cell.lower():
                            found_rows.append(row)
                            return found_rows
        else:
            raise ValueError("Unsupported file format")

        return found_rows

    def search_excel_files_with_text(self, fname_match, search_text, progress_callback=None, search_results_callback=None, include_csv=False):
        self.searching = True
        excel_files = self.search_excel_files(fname_match, progress_callback, include_csv)
        files_with_text = []

        for file in excel_files:
            if not self.searching:
                break
            progress_callback(str(file))
            found_rows = self.search_excel(file, search_text)
            if found_rows:
                files_with_text.append((file, found_rows))
                search_results_callback(os.path(file), "./", os.path.basename(file), found_rows)

        self.searching = False
        return files_with_text

    def stop_search(self):
        self.searching = False


class AppConfig:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config_file = os.path.join(tempfile.gettempdir(), 'app_config.ini')
        self.load_config()

    def load_config(self):
        self.config.read(self.config_file)

    def save_config(self, path, fname_match, search_text, open_in_editor, recursive_search, include_csv):
        if not self.config.has_section('LAST_INPUTS'):
            self.config.add_section('LAST_INPUTS')
        self.config.set('LAST_INPUTS', 'path', path)
        self.config.set('LAST_INPUTS', 'fname_match', fname_match)
        self.config.set('LAST_INPUTS', 'search_text', search_text)
        self.config.set('LAST_INPUTS', 'open_in_editor', str(open_in_editor))
        self.config.set('LAST_INPUTS', 'recursive_search', str(recursive_search))
        self.config.set('LAST_INPUTS', 'include_csv', str(include_csv))
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Searcher")
        self.config = AppConfig()
        self.searcher = None
        self.searching = False
        self.search_forced_stop = False
        self.width_fields = 80

        self.setup_ui()
        self.load_config()

        self.last_update_time = 0

    def setup_ui(self):
        # Path input
        self.label_path = tk.Label(self.root, text="Path to search:")
        self.label_path.grid(row=0, column=0, padx=10, pady=5, sticky='e')
        
        self.entry_path = tk.Entry(self.root, width=self.width_fields)
        self.entry_path.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        
        # Search text input
        self.label_search_text = tk.Label(self.root, text="Search text:")
        self.label_search_text.grid(row=1, column=0, padx=10, pady=5, sticky='e')
        
        self.entry_search_text = tk.Entry(self.root, width=self.width_fields)
        self.entry_search_text.grid(row=1, column=1, padx=10, pady=5, sticky='w')

        # Filename match input
        self.label_fname_match = tk.Label(self.root, text="Filename match:")
        self.label_fname_match.grid(row=2, column=0, padx=10, pady=5, sticky='e')
        
        self.entry_fname_match = tk.Entry(self.root, width=self.width_fields)
        self.entry_fname_match.grid(row=2, column=1, padx=10, pady=5, sticky='w')

        self.button_browse = tk.Button(self.root, text="Browse", command=self.browse_path)
        self.button_browse.grid(row=0, column=3, padx=10, pady=5)

        # Checkbox to open results in text editor
        self.var_open_in_editor = tk.BooleanVar()
        self.check_open_in_editor = tk.Checkbutton(self.root, text="Open results in text editor", variable=self.var_open_in_editor)
        self.check_open_in_editor.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky='w')
        
        # Checkbox to search recursively
        self.var_recursive_search = tk.BooleanVar()
        self.check_recursive_search = tk.Checkbutton(self.root, text="All subfolders", variable=self.var_recursive_search)
        self.check_recursive_search.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky='w')
        
        # Control buttons
        self.button_search = tk.Button(self.root, text="Search", command=self.start_search)
        self.button_search.grid(row=5, column=1, padx=10, pady=5, sticky='e')

        self.button_stop_search = tk.Button(self.root, text="Stop", command=self.stop_search, state=tk.DISABLED)
        self.button_stop_search.grid(row=5, column=2, padx=10, pady=5)

        self.button_close = tk.Button(self.root, text="Close", command=self.close_application)
        self.button_close.grid(row=5, column=3, padx=10, pady=5, sticky='w')
        
        # Results display
        self.text_results = tk.Text(self.root, width=80, height=20)
        self.text_results.grid(row=6, column=0, columnspan=5, padx=10, pady=10, sticky='nsew')

        # Add scrollbar to the text widget
        self.scrollbar = tk.Scrollbar(self.root, command=self.text_results.yview)
        self.text_results.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.grid(row=6, column=5, sticky='nsew')

        # Status label
        self.status_label = tk.Label(self.root, text="Status: Ready")
        self.status_label.grid(row=7, column=0, columnspan=5, padx=10, pady=5, sticky='w')

        # Configure grid weights to make the text box expandable
        self.root.grid_rowconfigure(6, weight=10)
        self.root.grid_columnconfigure(1, weight=10)

        # Checkbox to include CSV files
        self.var_include_csv = tk.BooleanVar()
        self.check_include_csv = tk.Checkbutton(self.root, text="Include CSV files", variable=self.var_include_csv)
        self.check_include_csv.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='w')

        self.root.bind('<Return>', lambda event: self.start_search())
        self.root.bind('<Escape>', lambda event: self.stop_search())

    def load_config(self):
        self.config.load_config()
        if self.config.config.has_section('LAST_INPUTS'):
            self.entry_path.insert(0, self.config.config.get('LAST_INPUTS', 'path', fallback=''))
            self.entry_fname_match.insert(0, self.config.config.get('LAST_INPUTS', 'fname_match', fallback=''))
            self.entry_search_text.insert(0, self.config.config.get('LAST_INPUTS', 'search_text', fallback=''))
            self.var_open_in_editor.set(self.config.config.getboolean('LAST_INPUTS', 'open_in_editor', fallback=False))
            self.var_recursive_search.set(self.config.config.getboolean('LAST_INPUTS', 'recursive_search', fallback=False))
            self.var_include_csv.set(self.config.config.getboolean('LAST_INPUTS', 'include_csv', fallback=False))
    
    def save_config(self):
        self.config.save_config(
            self.entry_path.get(),
            self.entry_fname_match.get(),
            self.entry_search_text.get(),
            self.var_open_in_editor.get(),
            self.var_recursive_search.get(),
            self.var_include_csv.get()
        )

    def browse_path(self):
        initial_dir = self.entry_path.get()
        if not os.path.isdir(initial_dir):
            initial_dir = "/"
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, folder_selected)

    def start_search(self):
        self.search_forced_stop = False
        self.searching = True
        self.button_search.config(state=tk.DISABLED)
        self.button_stop_search.config(state=tk.NORMAL)
        self.text_results.delete(1.0, tk.END)
        self.status_label.config(text="Status: Searching...")
        search_thread = threading.Thread(target=self.search_files)
        search_thread.daemon = True
        search_thread.start()

    def stop_search(self):
        if self.searcher:
            self.searcher.stop_search()
        self.searching = False
        self.button_search.config(state=tk.NORMAL)
        self.button_stop_search.config(state=tk.DISABLED)
        self.status_label.config(text="Status: Search stopped")

    def search_files(self):
        path = self.entry_path.get()
        fname_match = self.entry_fname_match.get()
        search_text = self.entry_search_text.get()
        open_in_editor = self.var_open_in_editor.get()
        recursive_search = self.var_recursive_search.get()
        include_csv = self.var_include_csv.get()

        if not path or not fname_match or not search_text:
            self.root.after(0, lambda: messagebox.showwarning("Input Error", "Please provide path, filename match, and search text."))
            self.searching = False
            self.root.after(0, lambda: self.button_search.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.button_stop_search.config(state=tk.DISABLED))
            return

        self.searcher = ExcelSearcher(path, recursive=recursive_search)
        found_files = self.searcher.search_excel_files_with_text(fname_match, search_text, self.update_progress, self.update_search_results, include_csv)

        if found_files:
            self.root.after(0, lambda: self.text_results.delete(1.0, tk.END))
            self.root.after(0, lambda: self.text_results.tag_config("link", foreground="blue", underline=True))
            self.root.after(0, lambda: self.text_results.tag_bind("link", "<Enter>", lambda e: self.text_results.config(cursor="hand2")))
            self.root.after(0, lambda: self.text_results.tag_bind("link", "<Leave>", lambda e: self.text_results.config(cursor="")))

            results = []

            for file, found_rows in found_files:
                if not self.searching:
                    break
                subdir_name = os.path.basename(os.path.dirname(file))
                file_name = os.path.basename(file)
                self.root.after(0, lambda fp=file, s=subdir_name, f=file_name, r=found_rows: self.update_search_results(fp, s, f, r))
                results.append((subdir_name, file_name, found_rows))

            if open_in_editor and self.searching:
                temp_file_path = self.write_results_to_temp_file(results)
                self.open_temp_file(temp_file_path)
        else:
            if not self.search_forced_stop:
                self.root.after(0, lambda: messagebox.showinfo("No Results", "No matching files found."))

        self.save_config()
        self.searching = False
        self.root.after(0, lambda: self.button_search.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.button_stop_search.config(state=tk.DISABLED))
        self.root.after(0, lambda: self.status_label.config(text="Status: Search done"))

    def update_progress(self, current_subdir):
        current_time = time.time()
        elapsed_time = current_time - self.last_update_time

        if elapsed_time >= 0.1:
            def update_text():
                self.status_label.config(text=f"Searching in: {PathShortener.shorten_path_pixels(current_subdir, widget=self.status_label)}")
                self.last_update_time = time.time()

            self.root.after(0, update_text)

    def update_search_results(self, full_path, subdir_name, file_name, found_rows):
        start_index = self.text_results.index(tk.INSERT)
        self.text_results.insert(tk.END, f"{subdir_name}/{file_name}\n")
        end_index = self.text_results.index(tk.INSERT)
        self.text_results.tag_add("link", start_index, end_index)
        self.text_results.tag_bind("link", "<Button-1>", lambda event, fp=full_path: self.open_file_location(fp))
        
        for row in found_rows:
            row_data = ', '.join([str(cell) for cell in row])
            self.text_results.insert(tk.END, f"    {row_data}\n")

    def write_results_to_temp_file(self, results):
        with tempfile.NamedTemporaryFile(delete=False, prefix="jentmp_", suffix='.txt', mode='w') as temp_file:
            for subdir_name, file_name, rows in results:
                temp_file.write(f"{subdir_name}/{file_name}\n")
                for row in rows:
                    row_data = ', '.join([str(cell) for cell in row])
                    temp_file.write(f"    {row_data}\n")
        return temp_file.name

    def open_temp_file(self, temp_file_path):
        if os.name == 'nt':
            os.startfile(temp_file_path)
        elif os.name == 'posix':
            subprocess.call(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', temp_file_path])
        else:
            print(f"Unsupported OS: {os.name}")

    def open_file_location(self, full_path):
        folder = os.path.dirname(full_path)
        if os.name == 'nt':
            os.startfile(folder)
        elif os.name == 'posix':
            subprocess.call(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', folder])
        else:
            print(f"Unsupported OS: {os.name}")

    def close_application(self):
        if self.searching:
            self.stop_search()
        self.root.destroy()
        os._exit(0)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
