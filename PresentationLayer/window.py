from BusinessLogicLayer.explorer_business_logic import get_drive_list, get_folder_list, get_file_list
from ttkbootstrap import Window, Menu, Text, Label, Entry, Button, Treeview, StringVar, Scrollbar, HORIZONTAL, VERTICAL
from ttkbootstrap.dialogs.dialogs import Messagebox
from tkinter import filedialog
from tkinter.ttk import Combobox
import os
import shutil
import fnmatch
import zipfile
from rarfile import RarFile
import patoolib


class MainWindow:
    def __init__(self):
        super().__init__()

        self.window = Window(title="File Explorer")

        self.window.grid_columnconfigure(0, weight=1, minsize=300)
        self.window.grid_columnconfigure(1, weight=1)
        self.window.grid_rowconfigure(1, weight=1)

        self.create_menu()

        self.current_path_variable = StringVar()
        self.path_entry = Entry(self.window, textvariable=self.current_path_variable, state="readonly")
        self.path_entry.grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

        self.search_entry = Entry(self.window, width=30)
        self.search_entry.grid(row=0, column=2, pady=10, padx=10, sticky="ew")

        self.search_button = Button(self.window, text="Search", command=self.search)
        self.search_button.grid(row=0, column=3, pady=10, padx=(0, 10), sticky="ew")

        self.theme_label = Label(self.window, text="Change Theme:")
        self.theme_label.grid(row=0, column=4, pady=10, padx=10, sticky="ew")

        self.theme_combobox = Combobox(self.window, width=15, state="readonly")
        self.theme_combobox['values'] = self.window.style.theme_names()
        self.theme_combobox.grid(row=0, column=5, pady=10, padx=(0, 10), sticky="ew")
        self.theme_combobox.current(2)
        self.window.bind("<<ComboboxSelected>>", self.change_theme)

        self.explorer_treeview = Treeview(self.window)
        self.explorer_treeview.insert("", "end", iid="ThisPC", text="This PC")
        self.explorer_treeview.grid(row=1, column=0, rowspan=2, pady=(0, 10), padx=10, sticky="nsew")

        self.explorer_table = Treeview(self.window, columns=("name", "type", "size",
                                                             "creation_date", "modified_date", "access_date"))
        self.explorer_table.grid(row=1, column=1, columnspan=5, pady=(0, 10), padx=(0, 10), sticky="nsew")

        self.explorer_table.heading("#0", text="NO", anchor="w")
        self.explorer_table.heading("#1", text="Name", anchor="w")
        self.explorer_table.heading("#2", text="Type", anchor="w")
        self.explorer_table.heading("#3", text="Size", anchor="w")
        self.explorer_table.heading("#4", text="Creation Date", anchor="w")
        self.explorer_table.heading("#5", text="Modified Date", anchor="w")
        self.explorer_table.heading("#6", text="Access Date", anchor="w")

        self.explorer_table.column("#0", width=100)
        self.explorer_table.column("#1", width=350)
        self.explorer_table.column("#2", width=100)
        self.explorer_table.column("#3", width=120)
        self.explorer_table.column("#4", width=210)
        self.explorer_table.column("#5", width=210)
        self.explorer_table.column("#6", width=210)

        self.drive_list = get_drive_list()
        for drive in self.drive_list:
            self.explorer_treeview.insert("ThisPC", "end", iid=drive, text=drive)

        self.row_list = []

        self.explorer_treeview.bind("<<TreeviewOpen>>", self.load_children)
        self.explorer_treeview.bind("<<TreeviewSelect>>", self.load_table)
        self.explorer_table.bind("<Button-3>", self.do_popup)

        self.window.mainloop()

    def load_children(self, event):
        entry_id = self.explorer_treeview.selection()[0]
        children = self.explorer_treeview.get_children(entry_id)
        for child in children:
            folder_list = get_folder_list(child)
            for entry in folder_list:
                self.explorer_treeview.insert(child, "end", iid=entry.full_path, text=entry.name)

    def load_table(self, event):
        for row in self.row_list:
            self.explorer_table.delete(row)
        self.row_list.clear()

        entry_id = self.explorer_treeview.selection()[0]
        self.current_path_variable.set(entry_id)
        folder_list = get_folder_list(entry_id)
        row_number = 1
        for folder in folder_list:
            row = self.explorer_table.insert("", "end", iid=folder.full_path, text=row_number,
                                             values=(folder.name, folder.entry_type, "",
                                                     folder.creation_date, folder.modified_date, folder.access_date))
            self.row_list.append(row)
            row_number += 1

        file_list = get_file_list(entry_id)
        for file in file_list:
            row = self.explorer_table.insert("", "end", iid=file.full_path, text=row_number,
                                             values=(file.name, file.entry_type, file.size,
                                                     file.creation_date, file.modified_date, file.access_date))
            self.row_list.append(row)
            row_number += 1

        self.manage_menu(self.row_list)

    def create_menu(self):
        self.menubar = Menu(self.window)

        self.menubar.add_command(label="New Folder", command=self.new_folder, state="disabled")
        self.menubar.add_command(label="Delete", command=self.delete, state="disabled")
        self.menubar.add_command(label="Rename", command=self.rename, state="disabled")
        self.menubar.add_command(label="Open ZIP file", command=self.open_zip_file, state="disabled")
        self.menubar.add_command(label="Extract ZIP file", command=self.extract_zip_file, state="disabled")
        self.menubar.add_command(label="Compress to ZIP file", command=self.compress_to_zip, state="disabled")
        self.menubar.add_command(label="Open RAR file", command=self.open_rar_file, state="disabled")
        self.menubar.add_command(label="Extract RAR file", command=self.extract_rar_file, state="disabled")
        self.menubar.add_command(label="Compress to RAR file", command=self.compress_to_rar, state="disabled")

        self.window.config(menu=self.menubar)

    def manage_menu(self, row_list):
        if len(row_list) > 0:
            for i in range(9):
                self.menubar.entryconfig(i, state="normal")
        else:
            for i in range(9):
                self.menubar.entryconfig(i, state="disabled")
            self.menubar.entryconfig(0, state="normal")

    def do_popup(self, event):
        try:
            self.menubar.tk_popup(event.x_root, event.y_root)
        finally:
            self.menubar.grab_release()

    def new_folder(self):
        new_folder_form = Window(title="New Folder")

        new_folder_label = Label(new_folder_form, text="New Folder Name:")
        new_folder_label.grid(row=0, column=0, pady=10, padx=10)

        new_folder_entry = Entry(new_folder_form, width=50)
        new_folder_entry.grid(row=0, column=1, pady=10, padx=10, sticky="ew")

        def submit():
            try:
                folder_path = self.path_entry.get()
                file_path = new_folder_entry.get()

                if os.path.exists(f"{folder_path}\\{file_path}"):
                    Messagebox.show_error(message="Folder already exists")
                else:
                    os.mkdir(f"{folder_path}\\{file_path}")
                    self.load_table(None)
            except FileNotFoundError:
                Messagebox.show_error(message="Please select a Folder")
            new_folder_form.destroy()

        submit_button = Button(new_folder_form, text="Submit", command=submit)
        submit_button.grid(row=1, column=0, pady=10, padx=10, sticky="e")

        def cancel():
            new_folder_form.destroy()

        cancel_button = Button(new_folder_form, text="Cancel", command=cancel)
        cancel_button.grid(row=1, column=1, pady=10, padx=10, sticky="w")

    def delete(self):
        try:
            if os.path.exists(self.explorer_table.selection()[0]):
                shutil.rmtree(self.explorer_table.selection()[0])
                self.load_table(None)
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def rename(self):
        try:
            if os.path.exists(self.explorer_table.selection()[0]):
                relpath = os.path.relpath(self.explorer_table.selection()[0], self.explorer_treeview.selection()[0])

                rename_form = Window(title="Rename")

                rename_label = Label(rename_form, text="Rename To:")
                rename_label.grid(row=0, column=0, pady=10, padx=10)

                rename_entry = Entry(rename_form, width=50)
                rename_entry.insert(0, relpath)
                rename_entry.grid(row=0, column=1, pady=10, padx=10, sticky="ew")

                def submit():
                    folder_path = self.path_entry.get()
                    source_path = self.explorer_table.selection()[0]
                    target_path = rename_entry.get()

                    os.rename(source_path, f"{folder_path}\\{target_path}")
                    rename_form.destroy()
                    self.load_table(None)

                submit_button = Button(rename_form, text="Submit", command=submit)
                submit_button.grid(row=1, column=0, pady=10, padx=10, sticky="e")

                def cancel():
                    rename_form.destroy()

                cancel_button = Button(rename_form, text="Cancel", command=cancel)
                cancel_button.grid(row=1, column=1, pady=10, padx=10, sticky="w")

        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def open_zip_file(self):
        try:
            if os.path.exists(self.explorer_table.selection()[0]):
                zip_path = self.explorer_table.selection()[0]
                if zip_path.endswith(".zip"):
                    open_zip_file_form = Window(title="Open ZIP File")

                    open_zip_file_form.grid_rowconfigure(0, weight=1)
                    open_zip_file_form.grid_columnconfigure(0, weight=1)

                    y_scrollbar = Scrollbar(open_zip_file_form, orient=VERTICAL)
                    y_scrollbar.grid(row=0, column=1, sticky="ns")

                    x_scrollbar = Scrollbar(open_zip_file_form, orient=HORIZONTAL)
                    x_scrollbar.grid(row=1, column=0, sticky="ew")

                    zip_text = Text(open_zip_file_form, wrap="none", yscrollcommand=y_scrollbar.set,
                                    xscrollcommand=x_scrollbar.set)
                    zip_text.grid(row=0, column=0, sticky="nsew")

                    y_scrollbar.config(command=zip_text.yview)
                    x_scrollbar.config(command=zip_text.xview)

                    with zipfile.ZipFile(zip_path, "r") as zip_file:
                        for name in zip_file.namelist():
                            zip_text.insert(1.0, f"{name}\n")
                    zip_text.config(state="disabled")
                else:
                    Messagebox.show_error(message="File is not a zip file")
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def extract_zip_file(self):
        try:
            if os.path.exists(self.explorer_table.selection()[0]):
                try:
                    folder_path = self.explorer_treeview.selection()[0]
                    file_path = self.explorer_table.selection()[0]
                    target_path = filedialog.askdirectory(initialdir=folder_path)
                    with zipfile.ZipFile(file_path, "r") as zip_file:
                        zip_file.extractall(target_path)
                except PermissionError:
                    pass
            else:
                Messagebox.show_error(message="File is not a zip file")
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def compress_to_zip(self):
        try:
            source_folder = self.explorer_table.selection()[0]
            parent_folder = self.explorer_treeview.selection()[0]
            target_folder = filedialog.askdirectory(initialdir=parent_folder)
            zip_path = os.path.relpath(source_folder, parent_folder)
            target_zip_file = os.path.join(target_folder, f"{zip_path}.zip")

            with zipfile.ZipFile(target_zip_file, mode="w") as zip_file:
                for main_folder, folders, files in os.walk(source_folder):
                    for file in files:
                        full_path = os.path.join(main_folder, file)
                        relative_path = os.path.relpath(full_path, source_folder)
                        zip_file.write(full_path, relative_path)
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def open_rar_file(self):
        try:
            if os.path.exists(self.explorer_table.selection()[0]):
                file_path = self.explorer_table.selection()[0]
                if file_path.endswith(".rar"):
                    open_rar_file_form = Window(title="Open RAR File")

                    open_rar_file_form.grid_rowconfigure(0, weight=1)
                    open_rar_file_form.grid_columnconfigure(0, weight=1)

                    y_scrollbar = Scrollbar(open_rar_file_form, orient=VERTICAL)
                    y_scrollbar.grid(row=0, column=1, sticky="ns")

                    x_scrollbar = Scrollbar(open_rar_file_form, orient=HORIZONTAL)
                    x_scrollbar.grid(row=1, column=0, sticky="ew")

                    rar_text = Text(open_rar_file_form, wrap="none", yscrollcommand=y_scrollbar.set,
                                    xscrollcommand=x_scrollbar.set)
                    rar_text.grid(row=0, column=0, sticky="nsew")

                    y_scrollbar.config(command=rar_text.yview)
                    x_scrollbar.config(command=rar_text.xview)

                    with RarFile(file_path, "r") as rar_file:
                        for name in rar_file.namelist():
                            rar_text.insert(1.0, f"{name}\n")
                    rar_text.config(state="disabled")
                else:
                    Messagebox.show_error(message="File is not a rar file")
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def extract_rar_file(self):
        try:
            if os.path.exists(self.explorer_table.selection()[0]):
                try:
                    folder_path = self.explorer_treeview.selection()[0]
                    file_path = self.explorer_table.selection()[0]
                    target_path = filedialog.askdirectory(initialdir=folder_path)
                    with RarFile(file_path, "r") as rar_file:
                        rar_file.extractall(target_path)
                except PermissionError:
                    pass
            else:
                Messagebox.show_error(message="File is not a rar file")
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def compress_to_rar(self):
        try:
            source_folder = self.explorer_table.selection()[0]
            parent_folder = self.explorer_treeview.selection()[0]
            target_folder = filedialog.askdirectory(initialdir=parent_folder)
            zip_path = os.path.relpath(source_folder, parent_folder)
            target_zip_file = os.path.join(target_folder, f"{zip_path}.rar")
            patoolib.create_archive(archive=target_zip_file, filenames=(source_folder,))
        except IndexError:
            Messagebox.show_error(message="Please select a File or Folder")

    def change_theme(self, event):
        self.window.style.theme_use(self.theme_combobox.get())

    def search(self):
        if self.search_entry.get():
            try:
                root_path = self.explorer_treeview.selection()[0]
                term = self.search_entry.get()
                folder_list = []
                file_list = []
                print(f"Term: {term}")
                for main_folder, folders, files in os.walk(root_path):
                    for folder in folders:
                        if fnmatch.fnmatch(folder, term):
                            full_path = os.path.join(main_folder, folder)
                            folder_list.append(full_path)
                            print(full_path)
                    for file in files:
                        if fnmatch.fnmatch(file, term):
                            full_path = os.path.join(main_folder, file)
                            file_list.append(full_path)
                            print(full_path)
                print(f"Result: {len(folder_list)} Folders and {len(file_list)} Files")
            except IndexError:
                Messagebox.show_error(message="Please select a File or Folder")
        else:
            Messagebox.show_error(message="Please enter a term for search")
