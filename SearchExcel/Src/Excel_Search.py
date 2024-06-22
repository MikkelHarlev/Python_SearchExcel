import tkinter as tk
from tkinter import filedialog, messagebox
import os
import fnmatch
import openpyxl
import tempfile
import subprocess
import configparser
import threading
import time

class ExcelSearcher:
    def __init__(self, base_folder, recursive=False):
        self.base_folder = base_folder
        self.recursive = recursive
        self.searching = False

    def search_excel_files(self, fname_match, progress_callback=None):
        found_files = []
        if self.recursive:
            for root, _, files in os.walk(self.base_folder):
                if progress_callback:
                    progress_callback(root)  # Call the callback with the current subdir
                if not self.searching:
                    break
                for file_name in files:
                    if fnmatch.fnmatch(file_name, f'*{fname_match}*.xlsx'):
                        found_files.append(os.path.join(root, file_name))
        else:
            for subdir in os.listdir(self.base_folder):
                if not self.searching:
                    break
                subdir_path = os.path.join(self.base_folder, subdir)
                if os.path.isdir(subdir_path):
                    if progress_callback:
                        progress_callback(subdir_path)  # Call the callback with the current subdir
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
            if not self.searching:
                break
            for cell in row:
                if cell.value and search_text.lower() in str(cell.value).lower():
                    # If the text is found, append the entire row to found_rows
                    found_rows.append([cell.value for cell in row])
                    return found_rows  # Return the rows immediately if found

        return found_rows

    def search_excel_files_with_text(self, fname_match, search_text, progress_callback=None):
        self.searching = True
        excel_files = self.search_excel_files(fname_match, progress_callback)
        files_with_text = []

        for file in excel_files:
            if not self.searching:
                break
            found_rows = self.search_excel(file, search_text)
            if found_rows:
                files_with_text.append((file, found_rows))

        self.searching = False
        return files_with_text

    def stop_search(self):
        self.searching = False


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Searcher")

        # Initialize the config parser
        self.config = configparser.ConfigParser()
        self.config_file = 'app_config.ini'
        self.load_config()

        self.searching = False  # Flag to control the search process

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

        # Checkbox to search recursively
        self.var_recursive_search = tk.BooleanVar()
        self.check_recursive_search = tk.Checkbutton(root, text="Search recursively", variable=self.var_recursive_search)
        self.check_recursive_search.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
        
        self.var_recursive_search.set(self.config.getboolean('LAST_INPUTS', 'recursive_search', fallback=False))

        # Search button
        self.button_search = tk.Button(root, text="Search", command=self.start_search)
        self.button_search.grid(row=5, column=1, padx=10, pady=10)

        # Stop search button
        self.button_stop_search = tk.Button(root, text="Stop", command=self.stop_search, state=tk.DISABLED)
        self.button_stop_search.grid(row=5, column=2, padx=10, pady=10)
        
        # Results display
        self.text_results = tk.Text(root, width=80, height=20)
        self.text_results.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

        # Close button
        self.button_close = tk.Button(root, text="Close", command=self.close_application)
        self.button_close.grid(row=5, column=3, padx=10, pady=10)

        self.last_update_time = 0  # Variable to keep track of the last update time


    def load_config(self):
        self.config.read(self.config_file)
    
    def save_config(self):
        if not self.config.has_section('LAST_INPUTS'):
            self.config.add_section('LAST_INPUTS')
        self.config.set('LAST_INPUTS', 'path', self.entry_path.get())
        self.config.set('LAST_INPUTS', 'fname_match', self.entry_fname_match.get())
        self.config.set('LAST_INPUTS', 'search_text', self.entry_search_text.get())
        self.config.set('LAST_INPUTS', 'open_in_editor', str(self.var_open_in_editor.get()))
        self.config.set('LAST_INPUTS', 'recursive_search', str(self.var_recursive_search.get()))
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def browse_path(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, folder_selected)

    def start_search(self):
        self.searching = True
        self.button_search.config(state=tk.DISABLED)
        self.button_stop_search.config(state=tk.NORMAL)
        self.text_results.delete(1.0, tk.END)
        start_index = self.text_results.index(tk.INSERT)
        self.root.after(0, lambda si=start_index, fn=f"Searching:": "Dir: ")
        search_thread = threading.Thread(target=self.search_files)
        search_thread.daemon = True  # Make the thread a daemon thread
        search_thread.start()

    def stop_search(self):
        self.searcher.stop_search()  # stop_search the search in the searcher instance
        self.searching = False
        self.button_search.config(state=tk.NORMAL)
        self.button_stop_search.config(state=tk.DISABLED)

    def close_application(self):
        if self.searching:
            self.stop_search()  # Stop the search if ongoing
        self.root.destroy()  # Close the application
        os._exit(0)  # Forcefully terminate the program

    def search_files(self):
        path = self.entry_path.get()
        fname_match = self.entry_fname_match.get()
        search_text = self.entry_search_text.get()
        open_in_editor = self.var_open_in_editor.get()
        recursive_search = self.var_recursive_search.get()

        if not path or not fname_match or not search_text:
            self.root.after(0, lambda: messagebox.showwarning("Input Error", "Please provide path, filename match, and search text."))
            self.searching = False
            self.root.after(0, lambda: self.button_search.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.button_stop_search.config(state=tk.DISABLED))
            return

        self.searcher = ExcelSearcher(path, recursive=recursive_search)
        found_files = self.searcher.search_excel_files_with_text(fname_match, search_text, self.update_progress)

        if found_files:
            self.root.after(0, lambda: self.text_results.delete(1.0, tk.END))
            self.root.after(0, lambda: self.text_results.tag_config("link", foreground="blue", underline=True))
            self.root.after(0, lambda: self.text_results.tag_bind("link", "<Button-1>", self.open_file_location))
            results = []  # Store results for text editor

            for file, found_rows in found_files:
                if not self.searching:
                    break
                subdir_name = os.path.basename(os.path.dirname(file))
                file_name = os.path.basename(file)

                start_index = self.text_results.index(tk.INSERT)
                self.root.after(0, lambda si=start_index, fn=f"{subdir_name}/{file_name}\n": self.text_results.insert(si, fn))
                end_index = self.text_results.index(tk.INSERT)
                self.root.after(0, lambda si=start_index, ei=end_index: self.text_results.tag_add("link", si, ei))

                for row in found_rows:
                    row_data = ', '.join([str(cell) for cell in row])
                    self.root.after(0, lambda rd=f"    {row_data}\n": self.text_results.insert(tk.END, rd))

                results.append((subdir_name, file_name, found_rows))

            if open_in_editor and self.searching:
                temp_file_path = self.write_results_to_temp_file(results)
                self.open_temp_file(temp_file_path)
        else:
            self.root.after(0, lambda: messagebox.showinfo("No Results", "No matching files found."))

        # Save the current inputs
        self.save_config()

        # Reset search state
        self.searching = False
        self.root.after(0, lambda: self.button_search.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.button_stop_search.config(state=tk.DISABLED))

    def update_progress(self, current_subdir):
        current_time = time.time()  # Get the current time
        elapsed_time = current_time - self.last_update_time

        if elapsed_time >= 0.1:  # Check if at least 100ms have passed
            def update_text():
                self.text_results.delete("end-1c linestart", "end")
                self.text_results.insert(tk.END, f"Searching in: {current_subdir}")
                self.last_update_time = time.time()  # Update the last update time
            
            self.root.after(0, update_text)

    def write_results_to_temp_file(self, results):
        with tempfile.NamedTemporaryFile(delete=False, prefix="jentmp_", suffix='.txt', mode='w') as temp_file:
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
