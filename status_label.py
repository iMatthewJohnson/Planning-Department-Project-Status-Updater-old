import tkinter as tk

class StatusLabel():
    def __init__(self, parent, row, column, columnspan):

        self.parent = parent
        self.label = tk.Label(parent,text="",  justify=tk.LEFT, anchor='w')
        self.label.grid(row=row, column=column, columnspan=columnspan, sticky='w', padx=(25, 25), pady=(0, 15))

        parent.update_idletasks()
        wrap_length = parent.winfo_width() - self.label.winfo_x() - 30 # 20 is for some right padding
        print(wrap_length)
        self.label.config(wraplength=wrap_length)

    def error_message(self, text):
        self.label.config(text=text, fg='red')

    def success_message(self, text):
        self.label.config(text=text, fg='green')

    def reset_label(self):
        self.label.config(text='')
        self.parent.update_idletasks()