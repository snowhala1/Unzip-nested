# Unzip-nested

A simple drag-and-drop PowerShell GUI tool to recursively unzip nested archive files — even if the file extensions are non-standard. Supports password-protected archives and optional cleanup.

Created for personal use.

## Features

- 🖱️ Drag and drop one or more archive files (e.g., `.zip`, `.abc`, `.dat`)
- 📂 Extracts to a new folder named `<filename>_unzipped` in the same directory
- 🔁 Recursively extracts nested archive files
- 🔐 Supports password-protected archives (one password for all nested file)
- 🗑️ Option to delete original files after successful extraction to Recycle Bin
- 💡 Built with PowerShell and Windows Forms

## Requirements

- Windows
- [7-Zip](https://www.7-zip.org/) installed (default path: `C:\Program Files\7-Zip\7z.exe`)
- PowerShell 5.1+

## Usage

1. Download unzip-nested.exe.
2. Make sure `7z.exe` is installed and the path is correct in the script.
3. Run unzip-nested.exe
4. Drag the file you want to unzip to the interface.
