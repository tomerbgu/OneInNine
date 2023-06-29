import configparser
import os
import shutil
import sys
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.messagebox import showinfo

import ttkbootstrap as ttk
from tkinterdnd2 import *
from ttkbootstrap.constants import *
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.utility import enable_high_dpi_awareness

from model import Model


def create_directory(directory):
    if not os.path.exists(resource_path(directory)):
        os.makedirs(resource_path(directory))
        print("Directory created:", directory)
    else:
        print("Directory already exists:", directory)


def clear_directory(directory_path):
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            file_path = os.path.join(root, file)
            os.remove(file_path)
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            os.rmdir(dir_path)


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
        self.config_reader = configparser.ConfigParser()
        self.config_reader.read('config.ini')

        self.title('One in Nine')
        self.geometry("900x700")

        self.config(background="white")
        self.model = None
        self.dt = None

        style = ttk.Style()
        # Configure the style to increase the padding around the label
        style.configure("LabelSize.TLabel", padding=(0, 35))

        # Create a File Explorer label
        # labelframe = ttk.Label(self, text='dddd', bootstyle='secondary')
        self.label_file_explorer_calendar = ttk.Label(self, text="Calendar", background="#cdfeec", style="LabelSize.TLabel", width=120, anchor="center")
        button_explore_calendar = ttk.Button(self, text="Browse Files",
                                             command=lambda: self.browseFiles(self.label_file_explorer_calendar),
                                             bootstyle="success-outline")

        # labelframe2 = ttk.Labelframe(self, text='My Labelframe', bootstyle='secondary')
        self.label_file_explorer_data = ttk.Label(self, text="Data", background="#cdfeec", style="LabelSize.TLabel", width=120, anchor="center")
        button_explore_data = ttk.Button(self, text="Browse Files",
                                         command=lambda: self.browseFiles(self.label_file_explorer_data),
                                         bootstyle="success-outline")

        self.label_file_explorer_calendar.drop_target_register(DND_FILES)
        self.label_file_explorer_calendar.dnd_bind('<<Drop>>', self.on_drop)

        self.label_file_explorer_data.drop_target_register(DND_FILES)
        self.label_file_explorer_data.dnd_bind('<<Drop>>', self.on_drop)

        self.calendar_path = StringVar()
        self.data_path = StringVar()
        self.file_exists = self.check_if_files_exist()
        button_calc = ttk.Button(self, text="Calculate", command=self.calculate,
                                 bootstyle="success")
        button_exit = ttk.Button(self, text="Exit", command=sys.exit, bootstyle="danger")

        self.label_file_explorer_calendar.place(relx=0.5, rely=0.1, anchor='center')
        button_explore_calendar.place(relx=0.5, rely=0.2, anchor='center')
        self.label_file_explorer_data.place(relx=0.5, rely=0.3, anchor='center')
        button_explore_data.place(relx=0.5, rely=0.4, anchor='center')
        button_exit.place(relx=0.1, rely=0.9, anchor='sw')
        button_calc.place(relx=0.9, rely=0.9, anchor='se')

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
        self.file_exists = False

    def calculate(self, ):
        if self.calendar_path.get() == '':
            print("need to choose calendar")
        elif self.data_path.get() == '':
            print("need to choose data")
        else:
            print("copying files")
            create_directory("data")
            if not self.file_exists:
                clear_directory(resource_path("data"))
                copy_file_to_directory(self.calendar_path.get(),
                                       resource_path(f"data/{self.config_reader.get('files', 'calendar_file')}"))
                copy_file_to_directory(self.data_path.get(),
                                       resource_path(f"data/{self.config_reader.get('files', 'data_file')}"))
            if self.model == None:
                self.model = Model()
            print("calculating:")
            self.results = self.model.solve_model()
            if self.dt is None:
                self.create_table(self.results)
            else:
                self.update_table(self.results)

    def create_table(self, results):
        if results.empty:
            showinfo(title='No Results',
                     message="We're sorry, but we couldn't find a solution that meets all of the requirements")
            return
        colors = self.style.colors
        self.dt = Tableview(
            master=self,
            coldata=results.columns,
            rowdata=results.values,
            paginated=True,
            searchable=True,
            bootstyle=PRIMARY,
            stripecolor=(colors.light, None),
            autofit=True
        )
        self.dt.view.configure(selectmode='none')
        self.dt.view.bind('<Button-3>', lambda x: None)
        self.dt.view.bind('<ButtonRelease-1>', lambda x: self.select_row(self.dt.view))
        self.dt.view.selection_remove(self.dt.view.selection())

        self.dt.place(relx=0.5, rely=0.7, anchor='center')

        button_retry = ttk.Button(self, text="Try Again", command=self.show_confirmation_dialog)
        button_retry.place(relx=0.35, rely=.95, anchor='s')

        button_download = ttk.Button(self, text="Download", command=self.download_results)
        button_download.place(relx=0.65, rely=.95, anchor='s')

    def download_results(self):
        path = f"{self.config_reader.get('files', 'output_file')}.xlsx"
        self.results.to_excel(path, index=False)
        showinfo(title='Download Successful', message=f"Results can be found at {path}")

    def update_table(self, results):
        if results.empty:
            showinfo(title='No Results',
                     message="We're sorry, but we couldn't find a solution that meets all of the requirements")
            return
        self.dt.delete_rows()
        [self.dt.insert_row(values=row.tolist()) for index, row in results.iterrows()]
        self.dt.load_table_data(clear_filters=True)
        self.dt.view.selection_remove(self.dt.view.selection())

    def select_row(self, tree):
        tree.selection_toggle(tree.focus())

    def retry(self):
        selected_items = self.dt.view.selection()
        selected_rows = [{"org": val[0], "lec": val[1], "date": val[2], "slot": val[3]}
                         for val in (self.dt.view.item(item)["values"] for item in selected_items)]

        self.model.add_custom_constraints(selected_rows)
        self.results = self.model.solve_model()
        self.update_table(self.results)

    def show_confirmation_dialog(self):
        answer = messagebox.askyesno("Confirmation", "Are you sure you want to proceed?")
        if answer:
            self.retry()
        else:
            print("User clicked No")

    def on_drop(self, event):
        # Get the list of files dropped into the window
        if event.widget == self.label_file_explorer_calendar:
            self.label_file_explorer_calendar.configure(text="File Opened: " + event.data)
            self.calendar_path.set(event.data)
        elif event.widget == self.label_file_explorer_data:
            self.label_file_explorer_data.configure(text="File Opened: " + event.data)
            self.data_path.set(event.data)
        self.file_exists = False

    def check_if_files_exist(self):

        calendar_file = resource_path(f"data/{self.config_reader.get('files', 'calendar_file')}")
        data_file = resource_path(f"data/{self.config_reader.get('files', 'data_file')}")

        if os.path.exists(calendar_file) and os.path.exists(calendar_file):
            self.label_file_explorer_calendar.configure(text="File Opened: " + calendar_file)
            self.calendar_path.set(calendar_file)
            self.label_file_explorer_data.configure(text="File Opened: " + data_file)
            self.data_path.set(data_file)
            return True
        return False


# Let the window wait for any events
if __name__ == "__main__":
    window = PopupWindow()  # themename="superhero"
    enable_high_dpi_awareness(root=window, scaling=1)

    window.mainloop()
