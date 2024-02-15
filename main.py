import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import project_status_updater
from settings_manager import SettingsManager
from file_selection_widget import FileSelectionWidget
from notebook_widget import NotebookWidget
from status_label import StatusLabel
from settings_widget import SettingsWidget

def load_config():
    with open('config.json', 'r') as config_file:
        return json.load(config_file)

config = load_config()

SYNC_SUCCESS_MESSAGE = "Workbooks have been synced successfully."

COMMENTS_RESPONSE = 'comments_response'
ACTION_ID_RESPONSE = 'action_id_response'
STATUS_RESPONSE = 'status_response'

COMMENTS_PROJECT = 'comments_project'
ACTION_ID_PROJECT = 'action_id_project'
STATUS_A = 'status_a'
STATUS_B = 'status_b'
STATUS_C = 'status_c'
STATUS_D = 'status_d'
STATUS_E = 'status_e'
STATUS_F = 'status_f'


class ProjectStatusUpdaterApp(tk.Tk):  # Define the application class which inherits from tk.Tk
    APP_NAME = config['app_info']['title']
    APP_AUTHOR = config['app_info']['author']
    APP_WINDOW_DIMENSIONS = config['window']['dimensions']
    ERROR_COLOR = 'red'

    def __init__(self):
        super().__init__()

        # Bind the event of the window closing to the on_closing() function.
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.settings_manager = SettingsManager(self.APP_NAME, self.APP_AUTHOR)

        self.title(self.APP_NAME)
        self.geometry(self.APP_WINDOW_DIMENSIONS)
        self.resizable(False, False)  # Make the window not resizable

        # Define instance variables that will be defined later in the setup
        self.quit_button = None
        self.run_button = None
        self.status_label = None
        self.responses_wb_path_entry = None
        self.project_status_wb_path_entry = None
        self.settings_widget = None

        self._create_widgets()


    def run_program(self):
        project_status_wb_path = self.project_status_file_selection_widget.get_entry_content()
        status_responses_wb_path = self.status_responses_file_selection_widget.get_entry_content()

        self.status_label.reset_label()

        error_messages = self._validate_file_paths(project_status_wb_path, status_responses_wb_path)
        if error_messages:
            self.status_label.error_message(text='ERROR: ' + error_messages)
            return

        project_status_updater.run_sync(project_status_wb_path, status_responses_wb_path, self.settings_manager.get_settings())
        self.status_label.success_message('SUCCESS: ' + SYNC_SUCCESS_MESSAGE)

    def on_closing(self):
        self.save_settings()
        self.destroy()


    def _create_widgets(self):
        nb = NotebookWidget(parent=self)
        main_tab = nb.add_tab('Main')
        settings_tab = nb.add_tab('Settings')
        self._setup_file_selection_widgets(main_tab)
        self.status_label = StatusLabel(parent=self, row=3, column=0, columnspan=2)

        self._create_settings_widgets(settings_tab)
        self._setup_buttons(nb)

    def _setup_file_selection_widgets(self, frame):

        self.project_status_file_selection_widget = FileSelectionWidget(controller=self, parent=frame, row=0,
                                                                        label_text=config['labels']['file_selection_project_status'])

        self.status_responses_file_selection_widget = FileSelectionWidget(controller=self, parent=frame, row=2,
                                                                          label_text=config['labels']['file_selection_status_responses'])

    def _setup_buttons(self, nb):
        button_frame = self._create_buttons_frame()
        self._create_run_button(button_frame, nb)
        self._create_quit_button(button_frame, nb)

    def _create_buttons_frame(self):
        button_frame = tk.Frame(self)
        button_frame.place(anchor='se', relx=1.0, rely=1.0, height=40)
        for i in range(2):  # To center both buttons in their respective columns
            button_frame.columnconfigure(i, weight=1)
        return button_frame

    def _create_run_button(self, button_frame, nb):
        self.run_button = tk.Button(button_frame, text="Run", command=self.run_program)
        self.run_button.grid(row=0, column=0)

    def _create_quit_button(self, button_frame, nb):
        self.quit_button = tk.Button(button_frame, text="Cancel", command=self.quit)
        self.quit_button.bind("<Button-1>", lambda event: self.save_settings())
        self.quit_button.grid(row=0, column=1, padx=(0, 25))

    def _create_settings_widgets(self, frame):
        self.settings_widget = SettingsWidget(controller=self, frame=frame)


        # Create a list with the info about the groups we want to create, each tuple contains the title and position
        group_settings = [('Response Workbook Column Settings', 'left'),
                          ('Project Status Workbook Column Settings', 'right')]

        # Create a dictionary for specific inputs for each group
        label_entry_settings = {
            'Response Workbook Column Settings': [('Comments', COMMENTS_RESPONSE, 0),
                                                  ('Action ID', ACTION_ID_RESPONSE, 1),
                                                  ('Status', STATUS_RESPONSE, 2)],
            'Project Status Workbook Column Settings': [('Comments', COMMENTS_PROJECT, 0),
                                                        ('Action ID', ACTION_ID_PROJECT, 1),
                                                        ('Status A', STATUS_A, 0, 2, 40),
                                                        ('Status B', STATUS_B, 1, 2, 40),
                                                        ('Status C', STATUS_C, 2, 2, 40),
                                                        ('Status D', STATUS_D, 3, 2, 40),
                                                        ('Status E', STATUS_E, 4, 2, 40),
                                                        ('Status F', STATUS_F, 5, 2, 40)]
        }

        # Using a for loop to create the two groups
        for title, position in group_settings:
            group = self.settings_widget.create_settings_group(title, position)

            # Now create the labels and entry fields for the response workbook
            for setting in label_entry_settings[title]:
                self.settings_widget.create_settings_label_and_entry(group, *setting)

                # Load settings to populate the entries with the saved values (if available)
        self.load_settings()

    def load_settings(self):
        for key, entry_value in self.settings_manager.get_settings().items():
            self.settings_widget.set_entry_value(key, entry_value)

    def save_settings(self):
        self.settings_manager.save_settings(self.settings_widget.get_all_entry_values())

    def settings_value_updated(self, settings_key, value):
        self.settings_manager.update_setting(settings_key, value)

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

    def browse_button_pressed(self, file_dialogue):

        # Ask and get the file path by selecting it. Limit to Excel files
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        file_dialogue.set_entry_content(file_path)


if __name__ == "__main__":
    app = ProjectStatusUpdaterApp()
    app.mainloop()
