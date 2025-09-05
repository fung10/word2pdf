# 1. Program Description

## Program Name
Word Batch to PDF Converter

## Program Purpose
This program aims to provide a user-friendly graphical interface tool that allows users to batch convert multiple Microsoft Word documents (including .docx, .doc, .rtf, .odt, etc.) into PDF files. It is specifically designed for processing large volumes of documents and offers flexible file naming rules and conflict resolution mechanisms.

## Key Features
*   **Graphical User Interface (GUI):** Built with Tkinter for intuitive operation.
*   **Drag-and-Drop Functionality:** Supports dragging Word files or folders containing Word files directly into the program. Also supports dragging a folder to set the output path.
*   **Batch Processing:** Can process multiple Word files simultaneously.
*   **Multi-threaded Conversion:** The conversion process runs in separate threads, ensuring the GUI remains responsive during conversion.
*   **Custom Naming Rules:**
    *   **"Original Name":** Uses the original Word file name for the PDF file.
    *   **"Remove Square Brackets":** Automatically removes all square brackets `[]` and their contents from the Word file name, then uses the cleaned name for the PDF file.
*   **Output File Conflict Handling:** If a target PDF file already exists, the program automatically renames the new file (e.g., `Document (1).pdf`).
*   **Real-time Log and Status Display:** Shows conversion progress, success, failure, and other information at the bottom of the interface.
*   **Conversion Stop Function:** Allows users to stop the batch conversion at any time during the process.
*   **Conversion Summary Window:** After conversion is complete, a detailed summary window pops up, showing the conversion results for each file (success, failure, renamed due to conflict).
*   **Supports Various Word Formats:** Supports .docx, .docm, .doc, .dotx, .dotm, .dot, .rtf, .odt, etc.

## Technology Stack
*   **Programming Language:** Python 3.x
*   **GUI Framework:** Tkinter
*   **Drag-and-Drop Library:** `tkinterdnd2`
*   **Word COM Automation:** `pywin32` (for interacting with Microsoft Word application)
*   **Multi-threading:** Python `threading` module

## Operating System Requirements
*   **Windows Operating System** (due to dependency on `pywin32` and Microsoft Word COM automation).
*   **Microsoft Word installed** (recommended 2007 or later).

## Program Architecture
The program is divided into two main parts:
1.  `main.py`: Responsible for creating and managing the GUI, handling user interactions, and coordinating the conversion flow.
2.  `word_to_pdf_converter.py`: Contains the core conversion logic, including file naming rules, Word COM automation operations, multi-threading management, and result collection.

## Program Screenshot (Description)
As direct screenshots cannot be provided, here's a description of the program interface layout:

*   **Main Window Title:** "Word Batch to PDF Converter"
*   **Top Section:** "Word Files to Convert:" label, followed by a large table (Treeview) displaying the original names of added Word files and their previewed converted PDF names.
*   **File Operation Buttons:** "Add Word Files...", "Clear All", "Remove Selected" located below the table.
*   **Output Directory:** "Output PDF Directory:" label, an entry field showing the current output path, next to a "Select Directory..." button.
*   **Naming Rule:** "PDF Naming Rule:" label, next to a dropdown menu with options like "Remove Square Brackets" and "Original Name".
*   **Conversion Controls:** "Start Batch Conversion" (large blue button) and "Stop Conversion" (red button, enabled during conversion) located in the center of the window.
*   **Log Area:** "Conversion Log/Status:" label, followed by a scrolled text box displaying real-time conversion logs and status messages.
*   **Conversion Summary Window (Pop-up):** Appears after conversion, containing a table showing the original file, converted PDF name, status (Success, Failed, Conflict Renamed), and messages. The top section displays total files, success count, failed count, and conflict renamed count.
