# Archiver

Archiver is a desktop application for managing and archiving letters (PDFs) with a user-friendly graphical interface fully in arabic.

## Features

- **Viewing Letters:** Open and view PDF letters directly within the application.
- **Letter Auto Numbering:** Automatically assigns sequential numbers to new letters based on year and type.
- **Advanced Search:** Search for letters by number, date, adressee, or keywords.
- **Organized Storage:** Stores PDF files in folders structured by year and adressee for easy retrieval.
- **Export to Excel:** Export all letter data to an Excel file for reporting or backup.

## Usage

1. Run the application:
   ```sh
   python archiver.py
   ```
2. Use the interface to add, search, view, and manage your letters.
3. Export your data to Excel using the export feature in the main window.

## Building Executable

To build a standalone executable, use PyInstaller:

```sh
pyinstaller archiver.spec
```

## Database

The application uses an SQLite database (`archive.db`). If the database does not exist, it will be created automatically using the schema in `db.sql`.
