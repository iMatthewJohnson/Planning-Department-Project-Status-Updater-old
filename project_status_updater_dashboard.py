import tkinter as tk
from tkinter import filedialog, messagebox
import os
import project_status_updater


class ProjectStatusUpdaterApp(tk.Tk):  # Define the application class which inherits from tk.Tk

    ERROR_COLOR = 'red'  # Constant for the error message color
    PADX_VALUE = 10  # Constant padding in the x direction
    PADY_UPPER_VALUE = (10, 0)  # Constant upper padding in the y direction
    PADY_LOWER_VALUE = (0, 5)  # Constant lower padding in the y direction
    BUTTON_PADX_VALUE = (0, 20)  # Constant padding for buttons in the x direction

    def __init__(self):
        super().__init__()
        self.title('Town of Duxbury Planning Dept Project Status Sync Tool')
        self.geometry('650x250')  # Increase the width to create space
        self.resizable(False, False)  # Make the window not resizable

        self.create_widgets()


    def create_widgets(self):  # Method to create the various widgets in the GUI

        # Create file browser widgets for the two workbooks. Values for positioning and text are passed to the method
        self.project_status_entry, _ = self.create_file_selection_widgets(
            0,
            "Select Project Status Workbook:",
            'project',
            self.PADX_VALUE,
            self.PADY_UPPER_VALUE,
            self.PADY_LOWER_VALUE,
            self.BUTTON_PADX_VALUE
        )
        self.status_responses_entry, _ = self.create_file_selection_widgets(
            2,
            "Select Status Responses Workbook:",
            'responses',
            self.PADX_VALUE,
            self.PADY_UPPER_VALUE,
            self.PADY_LOWER_VALUE,
            self.BUTTON_PADX_VALUE
        )

        # Label for displaying error messages. The color and position of label are defined here
        self.error_label = tk.Label(self, text="", fg=self.ERROR_COLOR, anchor='w')
        self.error_label.grid(row=4, column=0, columnspan=2, sticky='w', padx=self.PADX_VALUE, pady=(20, 5))

        # Configure the grid so that the first column resizes with the window
        self.grid_columnconfigure(0, weight=1)

        # Define the Run and Cancel buttons. The commands called on button press and their positioning are specified
        tk.Button(self, text="Run", command=self.run_program).grid(row=5, column=0, padx=(0, 5), pady=30, sticky='e')
        tk.Button(self, text="Cancel", command=self.quit).grid(row=5, column=1, padx=self.BUTTON_PADX_VALUE, pady=30)


    def create_file_selection_widgets(self, row, label_text, browse_arg, padx_value, pady_upper_value, pady_lower_value,
                                      button_padx_value):
        # Label for file selection
        tk.Label(self, text=label_text).grid(row=row, column=0, sticky='w', padx=padx_value, pady=pady_upper_value)

        # Entry and Browse button for workbook
        entry = tk.Entry(self, width=58)
        entry.grid(row=row + 1, column=0, padx=padx_value, pady=pady_lower_value, sticky='ew')
        button = tk.Button(self, text="Browse", command=lambda: self.browse_file(browse_arg))
        button.grid(row=row + 1, column=1, padx=button_padx_value, pady=pady_lower_value)

        return entry, button


    def browse_file(self, file_type):
        # Creates a dialog box that allows the user to select a file.
        # The "Excel files" denotation allows only .xlsx files to be selected.
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        # If a file is selected (meaning the file dialog did not return an empty string),
        # the path of the file is inserted into a specific entry field, based on the file type.
        if file_path:
            # If the file type is 'project',
            # the project_status_entry field is cleared and the new path is inserted.
            if file_type == 'project':
                self.project_status_entry.delete(0, tk.END)
                self.project_status_entry.insert(0, file_path)
            # If the file type is 'responses',
            # the status_responses_entry field is cleared and the new path is inserted.
            elif file_type == 'responses':
                self.status_responses_entry.delete(0, tk.END)
                self.status_responses_entry.insert(0, file_path)


    def run_program(self):
        project_status_wb_path = self.project_status_entry.get()
        status_responses_wb_path = self.status_responses_entry.get()

        self._reset_error_label()

        error_messages = self._validate_file_paths(project_status_wb_path, status_responses_wb_path)
        if error_messages:
            self.error_label.config(text="ERROR: " + error_messages)
            return

        project_status_updater.run_sync(project_status_wb_path, status_responses_wb_path)
        messagebox.showinfo("Success", "Workbooks have been synced successfully.")

    def _reset_error_label(self):
        self.error_label.config(text="")
        self.update_idletasks()
        wrap_length = self.winfo_width() - self.error_label.winfo_x() - 20 # 20 is for some right padding
        self.error_label.config(wraplength=wrap_length)

    def _validate_file_paths(self, project_status_wb_path, status_responses_wb_path):
        error_messages = []

        if not project_status_wb_path:
            error_messages.append('Project Status Workbook is missing')
        elif not os.path.isfile(project_status_wb_path):
            error_messages.append('Project Status Workbook file does not exist')

        if not status_responses_wb_path:
            error_messages.append('Status Responses Workbook is missing')
        elif not os.path.isfile(status_responses_wb_path):
            error_messages.append('Status Responses Workbook file does not exist')

        if error_messages:
            return '; '.join(error_messages)
        return None


if __name__ == "__main__":
    app = ProjectStatusUpdaterApp()
    app.mainloop()
