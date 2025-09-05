# 2. Installation Guide

This program can be installed and run in two ways: from source code or using an executable bundled with PyInstaller.

## Prerequisites
1.  **Operating System:** Must be **Windows**.
2.  **Microsoft Word:** You must have **Microsoft Word** installed on your computer (recommended 2007 or later). This is essential for the program's core conversion functionality.

---

## Method 1: Install and Run from Source Code (Recommended for Developers)

1.  **Install Python 3.x:**
    If Python is not already installed on your system, download and install the latest version of Python 3.x from the [official Python website](https://www.python.org/downloads/windows/). Make sure to check "Add Python to PATH" during installation.

2.  **Download the Code:**
    Download or clone the project's source code to your local machine.
    ```bash
    git clone <repository_url> # If a Git repository exists
    # Or simply download the ZIP file and extract it
    ```

3.  **Create and Activate a Virtual Environment:**
    This is a good practice to isolate project dependencies.
    Open Command Prompt (CMD) or PowerShell and navigate to the project root directory.
    ```bash
    cd /path/to/your/project
    python -m venv venv
    # Windows:
    .\venv\Scripts\activate
    # macOS/Linux (if viewing code on these systems, but cannot run):
    # source venv/bin/activate
    ```

4.  **Install Dependencies:**
    Within the virtual environment, install all dependencies listed in `requirements.txt`.
    ```bash
    pip install -r requirements.txt
    ```
    This will install `pywin32` and `tkinterdnd2`.

5.  **Run `pywin32` Post-Installation Script:**
    `pywin32` requires an additional step to correctly configure COM objects.
    ```bash
    python -m post_install
    ```
    If you encounter permission issues, try running Command Prompt as an administrator.

6.  **Run the Program:**
    ```bash
    python main.py
    ```

---

## Method 2: Using PyInstaller Executable (Recommended for General Users)

1.  **Download the Executable:**
    Download the pre-packaged executable (e.g., `main.exe`) from the distribution channel (e.g., project's release page or shared folder).

2.  **Place the Executable:**
    Place the downloaded `main.exe` file in any folder you prefer.

3.  **Run the Program:**
    Double-click the `main.exe` file to launch the program.

## Notes for Testing on a 'Clean' Machine
*   **Always confirm that Microsoft Word is installed.** This is the most common runtime issue. Without Word, the program cannot launch the Word COM object for conversion.
*   **Test drag-and-drop functionality.** Ensure `tkinterdnd2` works correctly in the packaged executable.
*   **Test conversion of different Word formats.** Verify that .doc, .docx, .rtf, .odt, etc., convert correctly.
*   **Test naming rules and conflict handling.** Create some files with similar names or square brackets and test the output.
