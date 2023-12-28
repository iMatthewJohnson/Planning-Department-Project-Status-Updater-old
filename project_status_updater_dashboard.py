import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import project_status_updater
from settings_manager import SettingsManager


class ProjectStatusUpdaterApp(tk.Tk):  # Define the application class which inherits from tk.Tk
    APP_NAME = 'ToD Planning Dept Project Status Sync Tool'
    APP_AUTHOR = 'gov.duxbury-ma'
    APP_WINDOW_DIMENSIONS = '750x350'
    ERROR_COLOR = 'red'  # Constant for the error message color

    def __init__(self):
        super().__init__()

        # Bind the event of the window closing to the on_closing() function.
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Define instance variables
        self.quit_button = None
        self.run_button = None
        self.error_label = None
        self.status_responses_entry = None
        self.project_status_entry = None

        self.settings_manager = SettingsManager(self.APP_NAME, self.APP_AUTHOR)

        self.title(self.APP_NAME)
        self.geometry(self.APP_WINDOW_DIMENSIONS)
        self.resizable(False, False)  # Make the window not resizable

        self._create_widgets()


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

    def on_closing(self):
        self.settings_manager.save_settings()
        self.destroy()

    def _create_widgets(self):  # Method to create the various widgets in the GUI

        nb = ttk.Notebook(self)
        nb.grid(row=0, column=0, sticky="nsew", padx=10, pady=(20, 5))

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
        self.project_status_entry = self._create_file_selection_widgets(
            main_tab,  # pass the main tab as the parent here
            0,
            "Select Project Status Workbook:",
            'project',
        )
        self.status_responses_entry = self._create_file_selection_widgets(
            main_tab,  # pass the main tab as the parent here
            2,
            "Select Status Responses Workbook:",
            'responses',
        )

        # Label for displaying error messages. The color and position of label are defined here
        self.error_label = tk.Label(self, text="", fg=self.ERROR_COLOR, justify=tk.LEFT, anchor="w")
        self.error_label.grid(row=3, column=0, columnspan=2, sticky='w', padx=(25,25), pady=(0, 15))


        # Configure the grid so that the first column resizes with the window
        self.grid_columnconfigure(0, weight=1)

        # Create settings elements within the settings tab
        self._create_settings_widgets(settings_tab)

        # Create the buttons' frame
        button_frame = tk.Frame(self)
        button_frame.place(anchor='se', relx=1.0, rely=1.0, height=40)

        # Define the Run and Cancel buttons. Have the settings changed everytime a user clicks to change tabs, or
        # cancel or run is clicked.
        nb.bind("<<NotebookTabChanged>>", lambda event: self.settings_manager.save_settings())
        self.run_button = tk.Button(button_frame, text="Run", command=self.run_program)
        self.run_button.bind("<Button-1>", lambda event: self.settings_manager.save_settings())
        self.run_button.grid(row=0, column=0)
        self.quit_button = tk.Button(button_frame, text="Cancel", command=self.quit)
        self.quit_button.bind("<Button-1>", lambda event: self.settings_manager.save_settings())
        self.quit_button.grid(row=0, column=1, padx=(0, 25))

        button_frame.columnconfigure(0, weight=1)  # To center the "Run" button in its column
        button_frame.columnconfigure(1, weight=1)  # To center the "Cancel" button in its column

    def _create_settings_widgets(self, frame):

        # Create the two different groups. One for response workbook settings, and one for project status workbook settings.
        response_group = self._create_settings_group(frame, "Response Workbook Column Settings", 'left')
        project_group = self._create_settings_group(frame, "Project Status Workbook Column Settings", 'right')

        # Create the labels and entry fields for the response workbook
        self._create_settings_label_and_entry(response_group, 'Comments', 'Comments_response', 0)
        self._create_settings_label_and_entry(response_group, 'Action ID', 'Action_ID_response', 1)
        self._create_settings_label_and_entry(response_group, 'Status', 'Status_response', 2)

        # Create labels and entry fields for the project group
        self._create_settings_label_and_entry(project_group, 'Comments', 'Comments_project', 0)
        self._create_settings_label_and_entry(project_group, 'Action ID', 'Action_ID_project', 1)
        self._create_settings_label_and_entry(project_group, 'Status A', 'Status A', 0, 2, 40)
        self._create_settings_label_and_entry(project_group, 'Status B', 'Status B', 1, 2, 40)
        self._create_settings_label_and_entry(project_group, 'Status C', 'Status C', 2, 2, 40)
        self._create_settings_label_and_entry(project_group, 'Status D', 'Status D', 3, 2, 40)
        self._create_settings_label_and_entry(project_group, 'Status E', 'Status E', 4, 2, 40)

        # Load settings to populate the entries with the saved values (if available)
        self.settings_manager.load_settings()


    def _create_settings_group(self, frame, text, side):
        group = tk.LabelFrame(frame, text=text, padx=5, pady=5)
        group.pack(padx=10, pady=10, side=side, anchor='n')
        return group


    def _create_settings_label_and_entry(self, group, label_text, settings_key, row, column=0, padx=0):
        var = tk.StringVar(value='')
        tk.Label(group, text=f"'{label_text}' column: ").grid(row=row, column=column, sticky='w', padx=(padx, 0))
        entry = tk.Entry(group, textvariable=var, width=2)
        entry.grid(row=row, column=column + 1)
        self.settings_manager.add_setting(settings_key, entry)



    def _create_file_selection_widgets(self, frame, row, label_text, browse_arg):
        padx = 10
        pady_lower = (10, 0)
        pady_upper = (0, 5)
        button_padx = (0, 20)
        label_width = 58


        # Label for file selection
        tk.Label(frame, text=label_text).grid(row=row, column=0, sticky='nw', padx=padx, pady=pady_upper)

        # Entry and Browse button for workbook
        entry = tk.Entry(frame, width=label_width)
        entry.grid(row=row + 1, column=0, padx=padx, pady=pady_lower, sticky='new')
        button = tk.Button(frame, text="Browse", command=lambda: self.browse_file(browse_arg))
        button.grid(row=row + 1, column=1, padx=button_padx, pady=pady_lower, sticky='n')

        # Stop the label row from expanding
        frame.grid_rowconfigure(row, weight=0)
        # Allow the entry and button row to expand
        frame.grid_rowconfigure(row + 1, weight=1)

        return entry

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
