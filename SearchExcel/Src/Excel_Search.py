import os
import fnmatch
import openpyxl
import tempfile
import subprocess

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
        is_found = False
        for row in sheet.iter_rows():
            if is_found:
                break
            for cell in row:
                if cell.value and search_text.lower() in str(cell.value).lower():
                    # If the text is found, append the entire row to found_rows
                    found_rows.append([cell.value for cell in row])
                    is_found = True
                    break  # Exit the inner loop to avoid adding the same row multiple times

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

# Example usage
base_folder = r'c:\_MeAndUs\Programming\Python\GitHub\Python\Excel_Search\xStuff\Excels'
fname_match = 'example'
search_text = 'Great Britain'

searcher = ExcelSearcher(base_folder)
found_files = searcher.search_excel_files_with_text(fname_match, search_text)

# Write the found files to a temporary file and open it
searcher.write_and_open_file_list(found_files)
