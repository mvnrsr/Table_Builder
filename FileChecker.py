import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class FileChecker:
    @staticmethod
    def select_file():
        # Initialize Tkinter and hide the root window
        root = tk.Tk()
        root.withdraw()

        # Show the file dialog and capture the selected file name
        file_path = filedialog.askopenfilename(
            title="Select file",
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm")]
        )

        root.destroy()  # Close the Tkinter root window

        return file_path

    @staticmethod
    def alert_file_must_be_closed():
        # Show an info alert
        messagebox.showinfo("File Open Error", "Please close the Excel file before loading it into the syntax builder.")

# # Usage example
# file_checker = FileChecker()
# file_path = file_checker.select_file()

# # If a file was selected, proceed with the operation
# if file_path:
#     # Check if the file is closed, and if not, show an alert
#     # ... (Your logic here)
#     pass
# else:
#     # If the user canceled the file selection, handle the cancellation without showing an alert
#     print("File selection was canceled.")
