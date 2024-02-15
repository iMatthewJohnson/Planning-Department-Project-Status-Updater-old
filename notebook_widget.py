from tkinter import ttk

# Notebook Styles
_NOTEBOOK_ROW = 0
_NOTEBOOK_COL = 0
_NOTEBOOK_STICKY = "nsew" # Stretch out to all 4 corners of frame
_NOTEBOOK_PADX = 10
_NOTEBOOK_PADY = (20, 5)

class NotebookWidget(ttk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=_NOTEBOOK_ROW, column=_NOTEBOOK_COL,
                           sticky=_NOTEBOOK_STICKY, padx=_NOTEBOOK_PADX, pady=_NOTEBOOK_PADY)

    def add_tab(self, text):
        tab = ttk.Frame(self.notebook)
        tab.grid_columnconfigure(index=0, weight=1)
        tab.grid_rowconfigure(index=0, weight=1)
        self.notebook.add(tab, text=text)
        return tab
