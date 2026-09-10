# Student Result Management System

A Python-based GUI application designed to handle student marks, compute results automatically, and save data directly into an Excel spreadsheet.

## Features
* **Student Data Entry**: Input student personal details and marks for 5 subjects.
* **Auto Calculation**: Instantly computes total marks, percentage, and pass/fail status.
* **Excel Database**: Automatically creates and updates `student_results.xlsx` using OpenPyXL.
* **Search Functionality**: Quickly look up individual student records using their Roll Number.
* **Treeview Display**: View all stored student records in a structured, searchable table view.

## Technologies Used
* **Python 3**
* **Tkinter** (Graphical User Interface)
* **OpenPyXL** (Excel sheet manipulation)

## How to Run the Project Locally
1. Clone this repository or download the ZIP file.
2. Ensure Python 3 is installed on your computer.
3. Install the required dependency:
   ```bash
   pip install openpyxl
