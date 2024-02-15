import tkinter as tk

_LABEL_ANCHOR='w'
_LABEL_STICKY='w'
_LABEL_PADX= 25
_LABEL_PADY=(0, 15)
_LABEL_RIGHT_PADDING = 30

class StatusLabel():
    def __init__(self, parent, row, column, columnspan):

        self.parent = parent
        self.label = tk.Label(parent,text="",  justify=tk.LEFT, anchor='w')
        self.label.grid(row=row, column=column, columnspan=columnspan, 
                        sticky=_LABEL_STICKY, padx=_LABEL_PADX, pady=_LABEL_PADY)   
        
        # Make it so the label can wrap across multiple rows
        # update_idletasks() to make sure parent UI is full updated to get proper width
        parent.update_idletasks()
        # Make the length of each line be the parent's frame's width, minus the label's frame's width, minus some extra
        # to create padding on the right.
        wrap_length = parent.winfo_width() - self.label.winfo_x() - _LABEL_RIGHT_PADDING
        self.label.config(wraplength=wrap_length)

    def error_message(self, text):
        self.label.config(text=text, fg='red')

    def success_message(self, text):
        self.label.config(text=text, fg='green')

    def reset_label(self):
        self.label.config(text='')
        self.parent.update_idletasks()