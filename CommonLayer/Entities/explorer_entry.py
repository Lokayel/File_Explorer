from dataclasses import dataclass


@dataclass
class ExplorerEntry:
    name: str
    full_path: str
    entry_type: str
    size: str
    creation_date: str
    modified_date: str
    access_date: str