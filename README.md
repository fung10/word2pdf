# Word Batch to PDF Converter

## Project Overview

The Word Batch to PDF Converter is a user-friendly desktop application designed to efficiently convert multiple Microsoft Word documents (including .docx, .doc, .rtf, .odt, etc.) into PDF format. Built with Python's Tkinter, it provides a graphical interface that simplifies the batch conversion process, offering flexible file naming rules and robust error handling. This tool is ideal for users who frequently need to convert large numbers of Word files to PDF, such as administrative staff, researchers, or anyone managing document archives.

## Features

*   **Intuitive GUI:** A clean and easy-to-use graphical interface built with Tkinter.
*   **Drag-and-Drop Support:** Effortlessly add Word files or folders containing Word files by dragging them directly into the application. Output directories can also be set via drag-and-drop.
*   **Batch Processing:** Process multiple Word documents to PDF in a single operation.
*   **Multi-threaded Conversion:** The conversion process runs in separate threads, ensuring the GUI remains interactive during long processing tasks.
*   **Customizable Naming Rules:**
    *   **"Original Name":** Retain the original Word file name for the converted PDF.
    *   **"Remove Square Brackets":** Automatically clean file names by removing content within square brackets `[]` (e.g., `Document [Draft].docx` becomes `Document.pdf`).
*   **Output File Conflict Handling:** Automatically renames PDF files (e.g., `File (1).pdf`) if a file with the same name already exists in the output directory, preventing accidental overwrites.
*   **Real-time Logging:** Monitor the conversion process with live status updates, warnings, and error messages displayed directly in the application's log area.
*   **Stop Conversion:** Ability to halt an ongoing batch conversion process.
*   **Conversion Summary:** A detailed pop-up window after conversion, showing the status (success, failed, renamed due to conflict) for each processed file.
*   **Wide Format Support:** Converts various Word document formats including `.docx`, `.docm`, `.doc`, `.dotx`, `.dotm`, `.dot`, `.rtf`, and `.odt`.

## Prerequisites

*   **Operating System:** This application is designed for **Windows** only, as it relies on Microsoft Word COM automation (`pywin32`).
*   **Microsoft Word:** A functional installation of **Microsoft Word** (2007 or later recommended) is absolutely required on the machine where the application is run. Without Word, the program cannot perform conversions.

## Installation

This program is distributed as a pre-packaged executable for easy use.

1.  **Download the Executable:**
    Obtain the `main.exe` (or similarly named) executable from the project's release page or distribution source.

2.  **Place and Run:**
    Place the downloaded `main.exe` file in any folder you prefer and double-click it to launch the program.

## Usage

For a super quick start, see [QUICK_START.md](Converter_Handover/QUICK_START.md). For detailed steps:

1.  **Launch the application.**
    *   Double-click `main.exe`.

2.  **Add Word Files:**
    *   Click the "Add Word Files..." button and select one or more Word files.
    *   Alternatively, drag Word files or folders containing Word files directly into the large file list area of the program interface. The program will automatically scan folders for Word files and skip non-Word files.

3.  **Set Output PDF Directory:**
    *   Click the "Select Directory..." button and choose a folder for PDF output.
    *   Alternatively, drag a folder directly into the "Output PDF Directory:" entry field or onto the "Select Directory..." button next to it.

4.  **Select PDF Naming Rule:**
    *   From the "PDF Naming Rule:" dropdown menu, select the desired naming rule:
        *   **"Original Name":** PDF files will use the original Word file names.
        *   **"Remove Square Brackets":** PDF file names will have all square brackets `[]` and their contents removed.

5.  **Start Conversion:**
    *   Click the "Start Batch Conversion" button.
    *   The program will check if the output directory exists; if not, it will attempt to create it.
    *   If PDF files with the same name already exist in the output directory, the program will prompt a warning asking whether to continue (continuing will automatically rename the new files, e.g., `File (1).pdf`).
    *   Once conversion starts, the "Start Batch Conversion" button will be disabled, and the "Stop Conversion" button will be enabled.

6.  **Monitor Progress:**
    *   The "Conversion Log/Status:" area at the bottom will display real-time conversion logs and status messages.

7.  **Stop Conversion:**
    *   During conversion, click the "Stop Conversion" button to attempt to halt the batch conversion. The program will wait for currently processing files to complete, then stop new conversion tasks.

8.  **View Conversion Summary:**
    *   After conversion is complete, a "Conversion Summary" window will automatically pop up, detailing the conversion results for each file.

## Troubleshooting

*   **"Could not launch Word Application" error:** Ensure Microsoft Word is correctly installed and functional on your system.
*   **"File is currently in use or locked" error:** Ensure the Word file to be converted is not open in any other application (including Word itself).
*   **"Path too long" error:** Windows has a limit on file path length (typically around 255-260 characters). Try shortening the output directory path or the original file name.
*   **Program unresponsive:** While conversions run in a separate thread, extreme system load or specific COM errors might cause temporary unresponsiveness. The "Stop Conversion" button can be used to halt the process.
*   **Drag-and-Drop issues:** Ensure `tkinterdnd2` is correctly installed and that you are dragging valid files/folders.

## Technology Stack & Architecture

*   **Programming Language:** Python 3.x
*   **GUI Framework:** Tkinter
*   **Drag-and-Drop Library:** `tkinterdnd2`
*   **Word COM Automation:** `pywin32` (for interacting with Microsoft Word application)
*   **Multi-threading:** Python `threading` module

The program is divided into two main parts:
1.  `main.py`: Responsible for creating and managing the GUI, handling user interactions, and coordinating the conversion flow.
2.  `word_to_pdf_converter.py`: Contains the core conversion logic, including file naming rules, Word COM automation operations, multi-threading management, and result collection.

## Development

For detailed development information, including code structure, modification guides, testing procedures, and advanced tips, please refer to [DEVELOPMENT.md](Converter_Handover/DEVELOPMENT.md).

