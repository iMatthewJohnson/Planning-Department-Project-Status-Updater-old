import configparser
import os
import appdirs
import tkinter as tk

class SettingsManager:
    def __init__(self, app_name, app_author):
        self.settings_dir = appdirs.user_config_dir(app_name, app_author)
        self.settings_file = os.path.join(self.settings_dir, 'settings.ini')
        self.config = configparser.ConfigParser()
        self.settings_dict = {}  # Initialize an empty dictionary

        # Create directory if it doesn't exist
        if not os.path.exists(self.settings_dir):
            os.makedirs(self.settings_dir)

        # Load existing settings if they exist
        if os.path.exists(self.settings_file):
            self.config.read(self.settings_file)

    def add_setting(self, key, setting_val):
        self.settings_dict[key] = setting_val

    def save_settings(self):
        if not self.config.has_section('Settings'):  # Ensure 'Settings' section is present
            self.config.add_section('Settings')

        for setting_key, entry in self.settings_dict.items():
            self.config.set('Settings', setting_key, entry.get())

        with open(self.settings_file, 'w') as f:
            self.config.write(f)

    def load_settings(self):
        for setting_key, entry in self.settings_dict.items():
            value = self.config.get('Settings', setting_key, fallback='')
            entry.delete(0, tk.END)
            entry.insert(0, value)