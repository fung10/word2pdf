# 5. Development Guide

This section provides detailed development information for future maintenance and feature expansion.

## Development Tools and Environment
*   **Programming Language:** Python 3.x (recommended 3.8+)
*   **Integrated Development Environment (IDE):**
    *   Visual Studio Code (VS Code)
    *   PyCharm
    *   Or any other Python IDE you are comfortable with.
*   **Version Control:** Git (if the project uses Git for version control)
*   **Dependency Management:** `pip` and `venv` (Python virtual environment)
*   **Packaging Tool:** PyInstaller (for bundling Python programs into Windows executables)

## Software Requirements
*   **Operating System:** Windows 10/11 (or any Windows version supporting Microsoft Word COM automation).
*   **Python:** Python 3.x installed.
*   **Microsoft Word:** Microsoft Word must be installed (2007, 2010, 2013, 2016, 2019, 365 are all compatible).

## Project Dependencies
All dependencies are listed in `requirements.txt`.
```
pywin32==311
tkinterdnd2==0.4.3
```
Installation steps:
1.  `python -m venv venv`
2.  `.\venv\Scripts\activate` (Windows)
3.  `pip install -r requirements.txt`
4.  `python -m post_install` (for `pywin32` specific configuration)

## Code Structure
*   `main.py`: The main entry point for the Tkinter GUI application.
    *   `WordToPdfConverterApp` class: Handles all GUI elements, user interactions, event handling, and log display.
    *   Responsible for initializing `BatchConverter` and `WordConverterLogic`.
    *   Uses the `threading` module to start batch conversion in a separate thread to keep the GUI responsive.
*   `word_to_pdf_converter.py`: Core conversion logic module.
    *   `WordConverterLogic` class: Handles the naming logic for a single Word file (`get_pdf_filename`).
    *   `ConversionWorker` class: Inherits from `threading.Thread`. Each Worker thread launches an independent `win32com.client.DispatchEx("Word.Application")` instance to perform the actual Word to PDF conversion. This helps isolate COM errors and improves stability.
    *   `BatchConverter` class: Orchestrates multiple `ConversionWorker` threads. It manages the task queue (`queue.Queue`), results dictionary, shared filename tracker (`_shared_filename_tracker`), and locks (`threading.Lock`) to ensure correct filename conflict handling in a multi-threaded environment.
*   `requirements.txt`: List of Python dependencies.
*   `main.spec`: PyInstaller configuration file for packaging.

## Ways to Modify the Program

1.  **Add or Modify PDF Naming Rules:**
    *   **`main.py`:**
        *   Modify the `self.naming_rules` list to add new rule names.
        *   If the new rule requires additional user input, new GUI elements (e.g., entry fields, checkboxes) might need to be added.
    *   **`word_to_pdf_converter.py` (WordConverterLogic class):**
        *   In the `get_pdf_filename` method, add an `elif naming_rule == "Your New Rule":` block and implement the new naming logic.
        *   Ensure the logic is robust and handles various file name edge cases.

2.  **Adjust GUI Layout or Appearance:**
    *   All GUI-related modifications are done within the `WordToPdfConverterApp` class in `main.py`.
    *   Use Tkinter's `grid()` or `pack()` layout managers to adjust component positions.
    *   Modify `ttk.Style` or directly adjust colors, fonts, sizes, etc., in component configurations.
    *   Consider using `customtkinter` or other Tkinter extension libraries for a more modern look, but this will introduce new dependencies.

3.  **Enhance Error Handling and Logging:**
    *   **`_log` method:** In the `_log` methods in both `main.py` and `word_to_pdf_converter.py`, you can add functionality to write logs to a file for easier debugging of long-running or unattended conversion tasks.
    *   **COM Error Handling:** The `try-except pythoncom.com_error` block in `ConversionWorker` can be further refined to provide more specific advice or handling logic for particular COM error codes.
    *   **File Integrity Checks:** Preliminary checks for Word file integrity (e.g., if the file is corrupted) could be added before conversion, although this often requires more complex libraries or functionality from Word itself.

4.  **Support More File Types:**
    *   **`main.py` (add_word_files method):**
        *   Modify the `filetypes` list to add new file extensions.
        *   Update the `valid_extensions` tuple to include new extensions.
    *   **`word_to_pdf_converter.py`:**
        *   `win32com.client.DispatchEx("Word.Application")` typically handles all document types supported by Word. If support for Excel or PowerPoint is needed, different COM application instances (e.g., `Excel.Application` or `PowerPoint.Application`) would need to be launched, and their specific conversion methods implemented. This would be a significant feature expansion.

5.  **Adjust Multi-threading Behavior:**
    *   **`word_to_pdf_converter.py` (BatchConverter class):**
        *   Modify the `num_threads` parameter in the `convert_batch_threaded` method. Increasing the number of threads might speed up conversion but will also increase system resource consumption (each thread launches a Word instance) and could potentially lead to COM stability issues. It is recommended to test and adjust based on the target machine's CPU and RAM.

## Ways to Test the Program

1.  **Unit Testing:**
    *   Currently, there are no explicit unit tests in the program. It is recommended to write unit tests for the `get_pdf_filename` method in the `WordConverterLogic` class to ensure the correctness of the naming rule logic.
    *   Python's `unittest` or `pytest` frameworks can be used.

2.  **Manual Testing:**
    *   **File Addition:**
        *   Test adding a single Word file, multiple Word files.
        *   Test dragging and dropping a single file, multiple files, a single folder, multiple folders.
        *   Test adding a folder containing non-Word files, confirm non-Word files are correctly skipped.
        *   Test adding files already present in the list, confirm no duplicates are added.
    *   **File List Operations:**
        *   Test the "Clear All" button.
        *   Test selecting single or multiple files and using the "Remove Selected" button.
    *   **Output Directory:**
        *   Test selecting an existing directory.
        *   Test selecting a non-existent directory (the program should attempt to create it).
        *   Test dragging and dropping a directory.
        *   Test dragging and dropping a file into the output directory input box (should show an error).
    *   **Naming Rules:**
        *   Test the "Original Name" rule.
        *   Test the "Remove Square Brackets" rule with file names containing various bracket patterns (e.g., `[Draft] Document.docx`, `Document [Final].docx`, `[A][B]C.docx`).
    *   **Conversion Process:**
        *   Test small batch conversions (2-3 files).
        *   Test large batch conversions (10+ files).
        *   Test in-progress clicking the "Stop Conversion" button.
        *   Test converting corrupted or unopenable Word files.
        *   Test scenarios where PDF files with the same name already exist in the output directory during conversion, confirming automatic renaming.
        *   Test scenarios where Word files are locked by other programs during conversion.
        *   Test scenarios where the output path is too long.
    *   **Interface Responsiveness:**
        *   While conversion is in progress, try dragging the window, clicking other buttons, confirming the GUI remains responsive.
    *   **Summary Window:**
        *   Confirm the summary window correctly pops up after conversion and displays accurate counts for success, failure, renamed files, and a detailed list.
        *   Confirm that main window controls are restored to normal after closing the summary window.
    *   **Closing the Program:**
        *   Test closing the main window when no conversion is active.
        *   Test closing the main window while conversion is in progress (should prompt a warning).
        *   Test closing the main window while the summary window is open (should prompt a warning).

## Development Prompt (Hypothetical prompt used for program development)
"Please create a Python Tkinter GUI application for batch converting Word documents to PDF. The application should have the following features:
1.  **File Selection:** Support selecting multiple Word files via a file dialog, and dragging/dropping Word files or folders containing Word files into a list area.
2.  **Output Directory Selection:** Support selecting an output folder via a directory dialog, and dragging/dropping a folder into the output path entry field.
3.  **PDF Naming Rules:** Provide two naming rule options:
    *   'Original Name': Use the original Word file name.
    *   'Remove Square Brackets': Automatically remove all square brackets `[]` and their contents from the file name.
4.  **Batch Conversion:** Capable of processing all Word files in the list at once.
5.  **Multi-threaded Processing:** The conversion process should run in separate threads to ensure the GUI remains responsive during conversion. Each Word conversion should use an independent Word COM application instance for improved stability.
6.  **File Conflict Handling:** If a converted PDF file name conflicts with an existing file in the output directory, it should be automatically renamed (e.g., `Document (1).pdf`).
7.  **Real-time Logging:** Display conversion progress, success/failure messages at the bottom of the interface.
8.  **Stop Function:** Provide a button to allow users to stop the batch conversion while it's in progress.
9.  **Conversion Summary:** After conversion is complete, pop up a summary window showing the conversion results for each file (success, failure, renamed).
10. **Platform Restriction:** The application must run on Windows and leverage the `pywin32` library for Microsoft Word COM automation.
11. **Dependencies:** Ensure `requirements.txt` includes all necessary libraries."
