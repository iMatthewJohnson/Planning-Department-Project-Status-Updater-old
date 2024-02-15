import tkinter as tk

class SettingsWidget():
    def __init__(self, controller, frame):
        self.controller = controller
        self.frame = frame
        self.setting_entries = {}

    def create_settings_group(self, text, side):
        group = tk.LabelFrame(self.frame, text=text, padx=5, pady=5)
        group.pack(padx=10, pady=10, side=side, anchor='n')
        return group

    def create_settings_label_and_entry(self, group, label_text, settings_key, row, column=0, padx=0):
        var = tk.StringVar(value='')
        var.trace('w', lambda *args: self.controller.settings_value_updated(settings_key, var.get()))

        tk.Label(group, text=f"'{label_text}' column: ").grid(row=row, column=column, sticky='w', padx=(padx, 0))
        entry = tk.Entry(group, textvariable=var, width=2)
        entry.grid(row=row, column=column + 1)
        self.setting_entries[settings_key] = entry

    def get_all_entry_values(self):
        return {
            key: self.setting_entries[key].get()
            for key in self.setting_entries.keys()
        }

    def set_entry_value(self, key, value):
        if key in self.setting_entries:
            self.setting_entries[key].delete(0, tk.END)
            self.setting_entries[key].insert(0, value)

