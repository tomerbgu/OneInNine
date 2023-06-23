import os
import sys
from tkinter import *
from tkinter import filedialog
from tkinterdnd2 import *
import ttkbootstrap as ttk
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.constants import *
import shutil

from ttkbootstrap.utility import enable_high_dpi_awareness

import model


def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)
        print("Directory created:", directory)
    else:
        print("Directory already exists:", directory)

def copy_file_to_directory(source_file, destination_directory):
    shutil.copy(source_file, destination_directory)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class PopupWindow(ttk.Window, TkinterDnD.Tk):
    def __init__(self, **args):
        super().__init__(**args)

        self.title('One in Nine')
        self.geometry("500x500")

        self.config(background="white")

        # Create a File Explorer label
        # labelframe = ttk.Label(self, text='dddd', bootstyle='secondary')
        self.label_file_explorer_calendar = ttk.Label(self, text="Calendar")
        button_explore_calendar = ttk.Button(self, text="Browse Files", command=lambda: self.browseFiles(self.label_file_explorer_calendar), bootstyle="success-outline")

        # labelframe2 = ttk.Labelframe(self, text='My Labelframe', bootstyle='secondary')
        self.label_file_explorer_data = ttk.Label(self, text="Data")
        button_explore_data = ttk.Button(self, text="Browse Files", command=lambda: self.browseFiles(self.label_file_explorer_data), bootstyle="success-outline")
        button_calc = ttk.Button(self, text="Calculate", command=self.calculate, bootstyle="success")
        button_exit = ttk.Button(self, text="Exit", command=sys.exit, bootstyle="danger")

        # Grid method is chosen for placing
        # the widgets at respective positions
        # in a table like structure by
        # specifying rows and columns


        self.label_file_explorer_calendar.drop_target_register(DND_FILES)
        self.label_file_explorer_calendar.dnd_bind('<<Drop>>', self.on_drop)

        self.label_file_explorer_data.drop_target_register(DND_FILES)
        self.label_file_explorer_data.dnd_bind('<<Drop>>', self.on_drop)

        self.calendar_path = StringVar()
        self.data_path = StringVar()

        colors = self.style.colors

        coldata = [
            {"text": "LicenseNumber", "stretch": False},
            "CompanyName",
            {"text": "UserCount", "stretch": False},
        ]

        rowdata = [
            ('A123', 'IzzyCo', 12),
            ('A136', 'Kimdee Inc.', 45),
            ('A158', 'Farmadding Co.', 36)
        ]

        dt = Tableview(
            master=self,
            coldata=coldata,
            rowdata=rowdata,
            paginated=True,
            searchable=True,
            bootstyle=PRIMARY,
            stripecolor=(colors.light, None),
        )

        # labelframe.grid(column=2, row=1)
        # button_explore_calendar.grid(column=2, row=2)
        # labelframe2.grid(column=2, row=4)
        # button_explore_data.grid(column=2, row=5)
        # button_exit.grid(column=1, row=7)
        # button_calc.grid(column=3, row=7)
        # dt.grid(column=2, row=6)

        self.label_file_explorer_calendar.place(relx=0.5, rely=0.1, anchor='center')
        button_explore_calendar.place(relx=0.5, rely=0.2, anchor='center')
        self.label_file_explorer_data.place(relx=0.5, rely=0.3, anchor='center')
        button_explore_data.place(relx=0.5, rely=0.4, anchor='center')
        button_exit.place(relx=0.1, rely=0.9, anchor='sw')
        button_calc.place(relx=0.9, rely=0.9, anchor='se')
        dt.place(relx=0.5, rely=0.7, anchor='center')


    def browseFiles(self, label_file_explorer):
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Text files",
                                                          "*.xlsx*"),
                                                         ("all files",
                                                          "*.*")))

        # Change label contents
        label_file_explorer.configure(text="File Opened: " + filename)
        if label_file_explorer == self.label_file_explorer_calendar:
            self.calendar_path.set(filename)
        elif label_file_explorer == self.label_file_explorer_data:
            self.data_path.set(filename)

    def calculate(self):
        if self.calendar_path.get()=='':
            print("need to choose calendar")
        elif self.data_path.get()=='':
            print("need to choose data")
        else:
            print("copying files")
            create_directory("data")
            copy_file_to_directory(self.calendar_path.get(), resource_path("data"))
            copy_file_to_directory(self.data_path.get(), resource_path("data"))
            print("calcluating:")
            model.main()

    def on_drop(self, event):
        # Get the list of files dropped into the window
        if event.widget == self.label_file_explorer_calendar:
            self.label_file_explorer_calendar.configure(text="File Opened: " + event.data)
            self.calendar_path.set(event.data)
        elif event.widget == self.label_file_explorer_data:
            self.label_file_explorer_data.configure(text="File Opened: " + event.data)
            self.data_path.set(event.data)


# Let the window wait for any events
if __name__ == "__main__":
    window = PopupWindow()#themename="superhero"
    enable_high_dpi_awareness(root=window, scaling=1)

    window.mainloop()
