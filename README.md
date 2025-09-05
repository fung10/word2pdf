# Word Batch to PDF Converter

## Project Overview

The Word Batch to PDF Converter is a user-friendly desktop application designed to efficiently convert multiple Microsoft Word documents into PDF format. Built with Python's Tkinter, it provides a graphical interface that simplifies the batch conversion process, offering flexible naming conventions and robust error handling. This tool is ideal for users who frequently need to convert large numbers of Word files to PDF, such as administrative staff, researchers, or anyone managing document archives.

## Features

*   **Intuitive GUI:** A clean and easy-to-use graphical interface built with Tkinter.
*   **Drag-and-Drop Support:** Effortlessly add Word files or folders containing Word files by dragging them directly into the application. Output directories can also be set via drag-and-drop.
*   **Batch Conversion:** Process multiple Word documents to PDF in a single operation.
*   **Customizable Naming Rules:**
    *   **"Original Name":** Retain the original Word file name for the converted PDF.
    *   **"Remove Square Brackets":** Automatically clean file names by removing content within square brackets `[]` (e.g., `Document [Draft].docx` becomes `Document.pdf`).
*   **Output Conflict Handling:** Automatically renames PDF files (e.g., `File (1).pdf`) if a file with the same name already exists in the output directory, preventing accidental overwrites.
*   **Real-time Logging:** Monitor the conversion process with live status updates, warnings, and error messages displayed directly in the application's log area.
*   **Responsive UI:** Conversions run in a separate thread, ensuring the GUI remains interactive during long processing tasks.
*   **Stop Conversion:** Ability to halt an ongoing batch conversion process.
*   **Conversion Summary:** A detailed pop-up window after conversion, showing the status (success, failed, renamed due to conflict) for each processed file.
*   **Wide Format Support:** Converts various Word document formats including `.docx`, `.docm`, `.doc`, `.dotx`, `.dotm`, `.dot`, `.rtf`, and `.odt`.

## Prerequisites

*   **Operating System:** Windows (The application relies on Microsoft Word COM automation, which is Windows-specific).
*   **Microsoft Word:** A functional installation of Microsoft Word (2007 or later recommended) is required on the machine where the application is run.

## Installation

### Method 1: From Source (For Developers)

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd word-batch-to-pdf-converter # Or your project's root directory
    ```
2.  **Create and activate a virtual environment:**
    ```bash
    python -m venv venv
    .\venv\Scripts\activate
    ```
3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    (The `requirements.txt` should contain `pywin32` and `tkinterdnd2`).
4.  **Run `pywin32` post-installation script:**
    ```bash
    python -m post_install
    ```
    (Run as administrator if permission errors occur).
5.  **Run the application:**
    ```bash
    python main.py
    ```

### Method 2: Using Executable (For End-Users)

1.  **Download the executable:** Obtain the `main.exe` (or similarly named) executable from the project's release page or distribution source.
2.  **Run the application:** Double-click `main.exe`.

## Usage

1.  **Launch the application.**
2.  **Add Word Files:**
    *   Click "Add Word Files..." to select one or more Word documents.
    *   Alternatively, drag and drop Word files or folders containing Word files directly into the large file list area.
3.  **Select Output Directory:**
    *   Click "Select Directory..." next to "Output PDF Directory:" to choose where the converted PDFs will be saved.
    *   Alternatively, drag and drop a folder onto the output directory entry field.
4.  **Choose Naming Rule:**
    *   Select your preferred PDF naming convention from the "PDF Naming Rule:" dropdown: "Original Name" or "Remove Square Brackets".
5.  **Start Conversion:**
    *   Click the "Start Batch Conversion" button.
    *   Monitor the "Conversion Log/Status:" area for real-time updates.
6.  **Review Summary:** A "Conversion Summary" window will appear upon completion, detailing the results.

## Troubleshooting

*   **"Could not launch Word Application" error:** Ensure Microsoft Word is installed and functional on your system.
*   **"File is currently in use or locked" error:** Close any applications (including Word itself) that might be holding the source Word file open.
*   **"Path too long" error:** Windows has a path length limit. Try shortening the output directory path or the original Word file names.
*   **GUI unresponsive:** While conversions run in a separate thread, extreme system load or specific COM errors might cause temporary unresponsiveness. The "Stop Conversion" button can be used to halt the process.

## Development

The project is structured into `main.py` (GUI and orchestration) and `word_to_pdf_converter.py` (core conversion logic, multi-threading, COM interaction).

*   **Technology Stack:** Python 3.x, Tkinter, `tkinterdnd2`, `pywin32`.
*   **Extending Naming Rules:** New naming rules can be implemented in `word_to_pdf_converter.py`'s `get_pdf_filename` method and integrated into `main.py`'s GUI.
*   **Error Handling:** Further COM error handling can be refined in `ConversionWorker` for more specific diagnostics.

## License

[Specify your license here, e.g., MIT License]
