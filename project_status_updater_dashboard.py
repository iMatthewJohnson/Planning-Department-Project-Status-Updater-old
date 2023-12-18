import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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
        self.geometry('650x315')  # Increase the width to create space
        self.resizable(False, False)  # Make the window not resizable

        self.create_widgets()


    def create_widgets(self):  # Method to create the various widgets in the GUI

        nb = ttk.Notebook(self)
        nb.grid(row=0, column=0, sticky="nsew", padx=self.PADX_VALUE, pady=(20, 5))  # adding padding

        # Create two tabs
        main_tab = ttk.Frame(nb)
        settings_tab = ttk.Frame(nb)

        nb.add(main_tab, text='Main')
        nb.add(settings_tab, text='Settings')

        # Make the tabs stretchable
        main_tab.grid_columnconfigure(0, weight=1)
        main_tab.grid_rowconfigure(0, weight=1)
        settings_tab.grid_columnconfigure(0, weight=1)
        settings_tab.grid_rowconfigure(0, weight=1)

        # Create file browser widgets for the two workbooks. Values for positioning and text are passed to the method
        self.project_status_entry, _ = self.create_file_selection_widgets(
            main_tab,  # pass the main tab as the parent here
            0,
            "Select Project Status Workbook:",
            'project',
            self.PADX_VALUE,
            self.PADY_UPPER_VALUE,
            self.PADY_LOWER_VALUE,
            self.BUTTON_PADX_VALUE
        )
        self.status_responses_entry, _ = self.create_file_selection_widgets(
            main_tab,  # pass the main tab as the parent here
            2,
            "Select Status Responses Workbook:",
            'responses',
            self.PADX_VALUE,
            self.PADY_UPPER_VALUE,
            self.PADY_LOWER_VALUE,
            self.BUTTON_PADX_VALUE
        )

        # Label for displaying error messages. The color and position of label are defined here
        self.error_label = tk.Label(self, text="", fg=self.ERROR_COLOR, justify=tk.LEFT, anchor="w")
        self.error_label.grid(row=3, column=0, columnspan=2, sticky='w', padx=(25,25), pady=(0, 15))


    # Configure the grid so that the first column resizes with the window
        self.grid_columnconfigure(0, weight=1)

        # Create settings elements within the settings tab
        self.create_settings_widgets(settings_tab)

        # Create the buttons' frame
        buttonFrame = tk.Frame(self)
        buttonFrame.place(anchor='e', relx=1.0, y=285, height=40)  # adjust these values as needed

        # Define the Run and Cancel buttons
        tk.Button(buttonFrame, text="Run", command=self.run_program).grid(row=0, column=0)
        tk.Button(buttonFrame, text="Cancel", command=self.quit).grid(row=0, column=1, padx=(0, 25))  # padx will add padding to the right side

        buttonFrame.columnconfigure(0, weight=1)  # To center the "Run" button in its column
        buttonFrame.columnconfigure(1, weight=1)  # To center the "Cancel" button in its column


    def create_settings_widgets(self, frame):
        response_group = tk.LabelFrame(frame, text="Response Workbook Column Settings", padx=5, pady=5)
        response_group.pack(padx=10, pady=10, side='left')

        project_group = tk.LabelFrame(frame, text="Project Status Workbook Column Settings", padx=5, pady=5)
        project_group.pack(padx=10, pady=10, side='left')

        # setting variables
        self.comments_col = tk.StringVar(value='G')
        self.actionID_col = tk.StringVar(value='F')
        self.status_col = tk.StringVar(value='H')

        # Create labels and entry fields for the response group
        tk.Label(response_group, text="'Comments' column: ").grid(row=0, column=0)
        tk.Entry(response_group, textvariable=self.comments_col).grid(row=0, column=1)

        tk.Label(response_group, text="'Action ID' column: ").grid(row=1, column=0)
        tk.Entry(response_group, textvariable=self.actionID_col).grid(row=1, column=1)

        tk.Label(response_group, text="'Status' column: ").grid(row=2, column=0)
        tk.Entry(response_group, textvariable=self.status_col).grid(row=2, column=1)



    def create_file_selection_widgets(self, frame, row, label_text, browse_arg, padx_value, pady_upper_value, pady_lower_value,
                                      button_padx_value):
        # Label for file selection
        tk.Label(frame, text=label_text).grid(row=row, column=0, sticky='w', padx=padx_value, pady=pady_upper_value)

        # Entry and Browse button for workbook
        entry = tk.Entry(frame, width=58)
        entry.grid(row=row + 1, column=0, padx=padx_value, pady=pady_lower_value, sticky='ew')
        button = tk.Button(frame, text="Browse", command=lambda: self.browse_file(browse_arg))
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
        wrap_length = self.winfo_width() - self.error_label.winfo_x() - 30 # 20 is for some right padding
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
