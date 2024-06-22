import tkinter as tk
from tkinter import filedialog, messagebox
import os
import fnmatch
import openpyxl
import tempfile
import subprocess
import configparser


class ExcelSearcher:
    def __init__(self, base_folder):
        self.base_folder = base_folder

    def search_excel_files(self, fname_match):
        found_files = []
        for subdir in os.listdir(self.base_folder):
            subdir_path = os.path.join(self.base_folder, subdir)
            if os.path.isdir(subdir_path):
                for file_name in os.listdir(subdir_path):
                    if fnmatch.fnmatch(file_name, f'*{fname_match}*.xlsx'):
                        found_files.append(os.path.join(subdir_path, file_name))
        return found_files

    def search_excel(self, file_path, search_text):
        # Load the workbook and select the active worksheet
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # List to store the rows containing the found text
        found_rows = []

        # Iterate over each row to search for the text
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and search_text.lower() in str(cell.value).lower():
                    # If the text is found, append the entire row to found_rows
                    found_rows.append([cell.value for cell in row])
                    return found_rows  # Return the rows immediately if found
        return found_rows

    def search_excel_files_with_text(self, fname_match, search_text):
        excel_files = self.search_excel_files(fname_match)
        files_with_text = []

        for file in excel_files:
            found_rows = self.search_excel(file, search_text)
            if found_rows:
                files_with_text.append(file)

        return files_with_text

    def write_and_open_file_list(self, file_list):
        with tempfile.NamedTemporaryFile(delete=False, suffix='.txt', mode='w') as temp_file:
            for file in file_list:
                temp_file.write(file + '\n')
            temp_file_path = temp_file.name

        if os.name == 'nt':  # For Windows
            os.startfile(temp_file_path)
        elif os.name == 'posix':  # For macOS and Linux
            subprocess.call(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', temp_file_path])
        else:
            print(f"Unsupported OS: {os.name}")

# GUI with tkinter
class App:

    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Searcher")

        # Initialize the config parser
        self.config = configparser.ConfigParser()
        self.config_file = 'app_config.ini'
        self.load_config()

        # Path input
        self.label_path = tk.Label(root, text="Path to search:")
        self.label_path.grid(row=0, column=0, padx=10, pady=10)
        
        self.entry_path = tk.Entry(root, width=50)
        self.entry_path.grid(row=0, column=1, padx=10, pady=10)
        self.entry_path.insert(0, self.config.get('LAST_INPUTS', 'path', fallback=''))
        
        self.button_browse = tk.Button(root, text="Browse", command=self.browse_path)
        self.button_browse.grid(row=0, column=2, padx=10, pady=10)
        
        # Search text input
        self.label_search_text = tk.Label(root, text="Search text:")
        self.label_search_text.grid(row=1, column=0, padx=10, pady=10)
        
        self.entry_search_text = tk.Entry(root, width=50)
        self.entry_search_text.grid(row=1, column=1, padx=10, pady=10)
        self.entry_search_text.insert(0, self.config.get('LAST_INPUTS', 'search_text', fallback=''))

        # Filename match input
        self.label_fname_match = tk.Label(root, text="Filename match:")
        self.label_fname_match.grid(row=2, column=0, padx=10, pady=10)
        
        self.entry_fname_match = tk.Entry(root, width=50)
        self.entry_fname_match.grid(row=2, column=1, padx=10, pady=10)
        self.entry_fname_match.insert(0, self.config.get('LAST_INPUTS', 'fname_match', fallback=''))

        # Checkbox to open results in text editor
        self.var_open_in_editor = tk.BooleanVar()
        self.check_open_in_editor = tk.Checkbutton(root, text="Open results in text editor", variable=self.var_open_in_editor)
        self.check_open_in_editor.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
        
        self.var_open_in_editor.set(self.config.getboolean('LAST_INPUTS', 'open_in_editor', fallback=False))

        # Search button
        self.button_search = tk.Button(root, text="Search", command=self.search_files)
        self.button_search.grid(row=4, column=1, padx=10, pady=10)
        
        # Results display
        self.text_results = tk.Text(root, width=80, height=20)
        self.text_results.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

    def load_config(self):
        self.config.read(self.config_file)
    
    def save_config(self):
        if not self.config.has_section('LAST_INPUTS'):
            self.config.add_section('LAST_INPUTS')
        self.config.set('LAST_INPUTS', 'path', self.entry_path.get())
        self.config.set('LAST_INPUTS', 'fname_match', self.entry_fname_match.get())
        self.config.set('LAST_INPUTS', 'search_text', self.entry_search_text.get())
        self.config.set('LAST_INPUTS', 'open_in_editor', str(self.var_open_in_editor.get()))
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def browse_path(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.entry_path.insert(0, folder_selected)

    def search_files(self):
        path = self.entry_path.get()
        fname_match = self.entry_fname_match.get()
        search_text = self.entry_search_text.get()
        open_in_editor = self.var_open_in_editor.get()

        if not path or not fname_match or not search_text:
            messagebox.showwarning("Input Error", "Please provide path, filename match, and search text.")
            return

        searcher = ExcelSearcher(path)
        found_files = searcher.search_excel_files_with_text(fname_match, search_text)

        if found_files:
            self.text_results.delete(1.0, tk.END)
            self.text_results.tag_config("link", foreground="blue", underline=True)
            self.text_results.tag_bind("link", "<Button-1>", self.open_file_location)
            results = []  # Store results for text editor

            for file in found_files:
                subdir_name = os.path.basename(os.path.dirname(file))
                file_name = os.path.basename(file)
                found_rows = searcher.search_excel(file, search_text)

                start_index = self.text_results.index(tk.INSERT)
                self.text_results.insert(tk.END, f"{subdir_name}/{file_name}\n")
                end_index = self.text_results.index(tk.INSERT)
                self.text_results.tag_add("link", start_index, end_index)

                for row in found_rows:
                    row_data = ', '.join([str(cell) for cell in row])
                    self.text_results.insert(tk.END, f"    {row_data}\n")

                results.append((subdir_name, file_name, found_rows))

            if open_in_editor:
                temp_file_path = self.write_results_to_temp_file(results)
                self.open_temp_file(temp_file_path)
        else:
            messagebox.showinfo("No Results", "No matching files found.")

        # Save the current inputs
        self.save_config()

    def write_results_to_temp_file(self, results):
        with tempfile.NamedTemporaryFile(delete=False, suffix='.txt', mode='w') as temp_file:
            for subdir_name, file_name, rows in results:
                temp_file.write(f"{subdir_name}/{file_name}\n")
                for row in rows:
                    row_data = ', '.join([str(cell) for cell in row])
                    temp_file.write(f"    {row_data}\n")
        return temp_file.name

    def open_temp_file(self, temp_file_path):
        if os.name == 'nt':  # For Windows
            os.startfile(temp_file_path)
        elif os.name == 'posix':  # For macOS and Linux
            subprocess.call(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', temp_file_path])
        else:
            print(f"Unsupported OS: {os.name}")

    def open_file_location(self, event):
        index = self.text_results.index("@%s,%s" % (event.x, event.y))
        line = self.text_results.get(index + " linestart", index + " lineend")
        folder_path = line.split("/")[0]
        file_name = line.split("/")[1]
        full_path = os.path.join(self.entry_path.get(), folder_path, file_name)
        folder = os.path.dirname(full_path)

        if os.name == 'nt':  # For Windows
            os.startfile(folder)
        elif os.name == 'posix':  # For macOS and Linux
            subprocess.call(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', folder])
        else:
            print(f"Unsupported OS: {os.name}")


# Run the app
root = tk.Tk()
app = App(root)
root.mainloop()


# # Example usage
# base_folder = r'c:\_MeAndUs\Programming\Python\GitHub\Python\Excel_Search\xStuff\Excels'
# fname_match = 'example'
# search_text = 'Great Britain'

# searcher = ExcelSearcher(base_folder)
# found_files = searcher.search_excel_files_with_text(fname_match, search_text)

# # Write the found files to a temporary file and open it
# searcher.write_and_open_file_list(found_files)
