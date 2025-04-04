import math
import os
from CommonLayer.Entities.explorer_entry import ExplorerEntry
from datetime import datetime


def get_drive_list():
    drive_list = []
    drive_names = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    for drive_name in drive_names:
        name = f"{drive_name}:\\"
        if os.path.exists(name):
            drive_list.append(name)

    return drive_list


def get_folder_list(parent_path):
    folder_list = []
    try:
        for entry in os.scandir(parent_path):
            if entry.is_dir():
                entry_information = entry.stat()
                creation_date = str(datetime.fromtimestamp(int(entry_information.st_birthtime)))
                modified_date = str(datetime.fromtimestamp(int(entry_information.st_mtime)))
                access_date = str(datetime.fromtimestamp(int(entry_information.st_atime)))
                explorer_entry = ExplorerEntry(entry.name, entry.path, "File folder", "",
                                               creation_date, modified_date, access_date)
                folder_list.append(explorer_entry)
    except FileNotFoundError:
        pass
    except PermissionError:
        pass
    return folder_list


def get_file_list(parent_path):
    file_list = []
    try:
        for entry in os.scandir(parent_path):
            if entry.is_file():
                entry_information = entry.stat()
                size = int(math.ceil(entry_information.st_size / 1024))
                creation_date = str(datetime.fromtimestamp(int(entry_information.st_birthtime)))
                modified_date = str(datetime.fromtimestamp(int(entry_information.st_mtime)))
                access_date = str(datetime.fromtimestamp(int(entry_information.st_atime)))
                explorer_entry = ExplorerEntry(entry.name, entry.path, "File", f"{size:,} KB",
                                               creation_date, modified_date, access_date)
                file_list.append(explorer_entry)
    except FileNotFoundError:
        pass
    except PermissionError:
        pass
    return file_list
