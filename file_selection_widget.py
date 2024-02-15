import tkinter as tk
from tkinter import filedialog, ttk

# Entry styles
_ENTRY_WIDTH = 58
_ENTRY_PADX = 10
_ENTRY_PADY = (10, 0)
_ENTRY_COLUMN = 0
_ENTRY_STICKY = 'new' # Anchored at top and stretched width on left and right

# Button Styles
_BUTTON_PADX = (0, 20)
_BUTTON_PADY = (10, 0)
_BUTTON_COLUMN = 1
_BUTTON_STICKY = 'n' # Anchored at top
_BROWSE_BUTTON_TEXT = "Browse..."

# Label Styles
_LABEL_PADX = 10
_LABEL_PADY = (0, 5)
_LABEL_COLUMN = 0
_LABEL_STICKY = 'nw' # Anchored in upper left corner

class FileSelectionWidget(ttk.Frame):
    def __init__(self, controller, parent, row, label_text, **kwargs):
        super().__init__(parent, **kwargs)

        self.controller = controller

        # Label for file selection
        self.label = ttk.Label(parent, text=label_text)
        self.label.grid(sticky=_LABEL_STICKY, row=row, column=_LABEL_COLUMN,
                        padx=_LABEL_PADX, pady=_LABEL_PADY)

        # Entry for the file
        self.entry = ttk.Entry(parent, width=_ENTRY_WIDTH)
        self.entry.grid(sticky=_ENTRY_STICKY, row=row + 1, column = _ENTRY_COLUMN, padx=_ENTRY_PADX, pady=_ENTRY_PADY)

        # Browse button
        self.button = ttk.Button(parent, text=_BROWSE_BUTTON_TEXT, command=self.browse_button_pressed)
        self.button.grid(row=row + 1, column=_BUTTON_COLUMN, padx=_BUTTON_PADX, pady=_BUTTON_PADY,sticky=_BUTTON_STICKY)

        # Stop the label row from expanding
        parent.grid_rowconfigure(row, weight=0)
        # Allow the entry and button row to expand
        parent.grid_rowconfigure(row + 1, weight=1)

    def get_entry_content(self):
        return self.entry.get()

    def set_entry_content(self, content):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, content)

    def browse_button_pressed(self):
        self.controller.browse_button_pressed(self)