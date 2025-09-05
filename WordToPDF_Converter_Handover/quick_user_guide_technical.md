# 3. Quick User Guide for Technical Staff

1.  **Start the Program:**
    *   From source: In the project root directory, activate the virtual environment and run `python main.py`.
    *   From executable: Double-click `main.exe`.

2.  **Add Word Files:**
    *   Click the "Add Word Files..." button and select one or more Word files.
    *   Drag Word files or folders containing Word files directly into the file list area of the program interface. The program will automatically scan folders for Word files.

3.  **Set Output PDF Directory:**
    *   Click the "Select Directory..." button and choose a folder for PDF output.
    *   Drag a folder directly into the "Output PDF Directory:" entry field or onto the "Select Directory..." button next to it.

4.  **Select PDF Naming Rule:**
    *   From the "PDF Naming Rule:" dropdown menu, select the desired naming rule:
        *   "Original Name": PDF files will use the original Word file names.
        *   "Remove Square Brackets": PDF file names will have all square brackets `[]` and their contents removed.

5.  **Start Conversion:**
    *   Click the "Start Batch Conversion" button.
    *   The program will check if the output directory exists; if not, it will attempt to create it.
    *   If PDF files with the same name already exist in the output directory, the program will prompt a warning asking whether to continue (continuing will automatically rename the new files).
    *   Once conversion starts, the "Start Batch Conversion" button will be disabled, and the "Stop Conversion" button will be enabled.

6.  **Monitor Progress:**
    *   The "Conversion Log/Status:" area at the bottom will display real-time conversion logs and status messages.

7.  **Stop Conversion:**
    *   During conversion, click the "Stop Conversion" button to attempt to halt the batch conversion. The program will wait for currently processing files to complete, then stop new conversion tasks.

8.  **View Conversion Summary:**
    *   After conversion is complete, a "Conversion Summary" window will automatically pop up, detailing the conversion results for each file.

9.  **Troubleshooting:**
    *   **"Could not launch Word Application" error:** Check if Microsoft Word is correctly installed.
    *   **"File is currently in use or locked" error:** Ensure the Word file to be converted is not open in any other application (including Word itself).
    *   **"Path too long" error:** Windows has a limit on file path length (typically around 255-260 characters). Try shortening the output directory path or the original file name.
    *   **Program unresponsive:** Although the program uses multi-threading, in extreme cases (e.g., Word COM object hanging), it might temporarily become unresponsive. Try using "Stop Conversion" or force-closing the program.
