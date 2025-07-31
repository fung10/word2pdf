# docx_converter_logic.py
import os
import win32com.client
import pythoncom # Still needed for pythoncom.com_error and potentially other COM interactions
import re

class DocxConverterLogic:
    """
    Handles the core logic for converting DOCX files to PDF using MS Word COM automation.
    Designed to be used by a GUI or other application, providing logging via a callback.
    """
    def __init__(self, log_callback=None):
        """
        Initializes the converter logic.

        Args:
            log_callback (callable, optional): A function to call for logging messages.
                                               It should accept (message, tag) arguments.
                                               Defaults to None, in which case messages are printed to console.
        """
        self._log_callback = log_callback

    def _log(self, message, tag=None):
        """
        Internal logging method that uses the provided log_callback.
        If no callback is provided, it prints to console.
        """
        if self._log_callback:
            self._log_callback(message, tag)
        else:
            color_map = {
                "blue": "\033[94m",
                "orange": "\033[93m",
                "green": "\033[92m",
                "red": "\033[91m",
                "reset": "\033[0m"
            }
            colored_message = f"{color_map.get(tag, '')}{message}{color_map['reset']}"
            print(colored_message)

    def get_pdf_filename(self, docx_path, naming_rule):
        """
        Determines the output PDF filename based on the DOCX path and selected naming rule.
        This method is public because the GUI needs to preview the PDF names.

        Args:
            docx_path (str): The full path to the source DOCX file.
            naming_rule (str): The selected naming rule (e.g., "Original Name", "Remove Square Brackets").

        Returns:
            str: The calculated filename for the output PDF.
        """
        base_name = os.path.splitext(os.path.basename(docx_path))[0]

        if naming_rule == "Original Name":
            return f"{base_name}.pdf"
        elif naming_rule == "Remove Square Brackets":
            # This rule specifically removes content within square brackets and the brackets themselves.
            # Use non-greedy matching (.*?) to handle multiple bracketed sections correctly.
            cleaned_base_name = re.sub(r'\[.*?\]', '', base_name)
            # Remove any resulting multiple spaces and trim leading/trailing spaces
            cleaned_base_name = re.sub(r'\s+', ' ', cleaned_base_name).strip()
            return f"{cleaned_base_name}.pdf"
        else:
            self._log(f"Warning: Unknown naming rule '{naming_rule}'. Using 'Original Name' as fallback.", "orange")
            return f"{base_name}.pdf"

    def convert_batch(self, docx_file_list, output_dir, naming_rule):
        """
        Performs batch conversion of DOCX files to PDF using MS Word COM automation.
        This method is designed to be run in a separate thread.

        Args:
            docx_file_list (list): A list of full paths to DOCX files.
            output_dir (str): The directory where converted PDF files will be saved.
            naming_rule (str): The rule to apply for naming the output PDF files.

        Returns:
            tuple: (converted_count, failed_count, total_files)
        """
        word = None
        converted_count = 0
        failed_count = 0
        total_files = len(docx_file_list)

        try:
            # Ensure the output directory exists. This check is also done in GUI,
            # but it's good for the logic to be robust.
            if not os.path.isdir(output_dir):
                try:
                    os.makedirs(output_dir)
                    self._log(f"Created output directory: {output_dir}", "blue")
                except Exception as e:
                    self._log(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                    return converted_count, failed_count, total_files

            # Initialize Word Application - ALWAYS launch a new, isolated instance using DispatchEx.
            # DispatchEx ensures a new instance is created, even if Word is already running.
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                self._log("Launched a new, isolated Word Application instance.", "blue")
            except Exception as e:
                self._log(f"Error: Could not launch a new, isolated Word Application instance. Please ensure MS Word is installed and not corrupted. Details: {e}", "red")
                return converted_count, failed_count, total_files

            for i, docx_path in enumerate(docx_file_list):
                current_file_name_display = os.path.basename(docx_path)
                self._log(f"[{i+1}/{total_files}] Converting: {current_file_name_display}", "orange")

                doc = None # Initialize doc to None for proper cleanup in case of early errors
                try:
                    if not os.path.exists(docx_path):
                        self._log(f"Skipping '{current_file_name_display}': Source file does not exist.", "red")
                        failed_count += 1
                        continue

                    pdf_file_name = self.get_pdf_filename(docx_path, naming_rule)
                    pdf_path = os.path.join(output_dir, pdf_file_name)

                    # Open Word document with specific options for robustness and user experience.
                    # ReadOnly=True ensures the original file is not modified.
                    # ConfirmConversions=False prevents conversion dialogs for unusual file types.
                    # AddToRecentFiles=False keeps the user's recent files list clean.
                    doc = word.Documents.Open(
                        os.path.abspath(docx_path),
                        ReadOnly=True,
                        ConfirmConversions=False,
                        AddToRecentFiles=False
                    )
                    doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 is the wdFormatPDF enum value for saving as PDF
                    doc.Close(False) # Close the document, False means don't save changes to the original DOCX
                    self._log(f"Successfully converted: {current_file_name_display} -> {pdf_file_name}", "green")
                    converted_count += 1

                except pythoncom.com_error as com_e:
                    # Handle COM-specific errors, often providing more detail for Word automation issues.
                    error_message = f"Conversion of '{current_file_name_display}' failed due to COM error: {com_e}"
                    # Attempt to get more COM error details if available.
                    # com_e.ex_info is a tuple containing detailed error info (source, description, helpfile, helpcontext, scode).
                    # ex_info[1] (description) or ex_info[4] (scode) are often most useful.
                    if hasattr(com_e, 'ex_info') and com_e.ex_info and len(com_e.ex_info) > 1:
                        com_error_description = com_e.ex_info[1]
                        com_error_scode = com_e.ex_info[4]
                        error_message += f"\nDetails: {com_error_description} (HRESULT: {hex(com_error_scode)})"

                        # Check for common file locking error code (0x80070020 - ERROR_SHARING_VIOLATION).
                        # This HRESULT value is -2147024864 in signed decimal.
                        if com_error_scode == -2147024864:
                            error_message += "\nPossible cause: The file is currently in use or locked by another application (e.g., another Word instance). Please close the file and try again."
                    self._log(error_message, "red")
                    failed_count += 1
                    # Try to close the document even if save failed, to prevent Word from hanging.
                    if doc:
                        try:
                            doc.Close(False)
                        except Exception as close_e:
                            self._log(f"Error closing document after failed COM conversion: {close_e}", "red")

                except Exception as e:
                    error_message = f"Conversion of '{current_file_name_display}' failed: {e}"
                    self._log(error_message, "red")
                    failed_count += 1
                    # Try to close the document even if save failed, to prevent Word from hanging.
                    if doc:
                        try:
                            doc.Close(False)
                        except Exception as close_e:
                            self._log(f"Error closing document after failed general conversion: {close_e}", "red")

        except Exception as e:
            self._log(f"Fatal error during batch conversion: {e}", "red")
        finally:
            if word:
                try:
                    word.Quit() # Always quit the Word application instance we launched.
                    del word # Release COM object to prevent memory leaks and ensure clean shutdown.
                    self._log("Word Application quit and COM object released.", "blue")
                except Exception as e:
                    self._log(f"Error quitting Word application: {e}", "red")

        return converted_count, failed_count, total_files