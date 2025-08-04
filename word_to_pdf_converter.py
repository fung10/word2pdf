# word_to_pdf_converter.py
import os
import win32com.client
import pythoncom
import re
import sys # Import sys module for platform check

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
        Note: This method provides the *intended* filename before uniqueness resolution.

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
            # Corrected regex: matches literal brackets by escaping them.
            cleaned_base_name = re.sub(r'\[.*?\]', '', base_name)
            # Remove any resulting multiple spaces and trim leading/trailing spaces
            cleaned_base_name = re.sub(r'\s+', ' ', cleaned_base_name).strip()
            # Handle cases where cleaning leaves an empty string or just spaces
            if not cleaned_base_name:
                cleaned_base_name = "Untitled_Document" # Fallback name
            return f"{cleaned_base_name}.pdf"
        else:
            self._log(f"Warning: Unknown naming rule '{naming_rule}'. Using 'Original Name' as fallback.", "orange")
            return f"{base_name}.pdf"

    def _get_unique_pdf_path(self, output_dir, proposed_pdf_filename, generated_filenames_tracker):
        """
        Generates a unique PDF path by appending a counter if a file with the
        same name already exists on disk or has been generated in this batch.

        Args:
            output_dir (str): The directory where the PDF will be saved.
            proposed_pdf_filename (str): The initial proposed filename (e.g., "document.pdf").
            generated_filenames_tracker (dict): A dictionary mapping base filenames
                                                to the *next counter to try* for that base.

        Returns:
            str: A unique full path for the PDF file.
        """
        base_name, ext = os.path.splitext(proposed_pdf_filename)
        
        # Get the starting counter for this base_name. If not seen, start from 0 (meaning no counter suffix yet).
        # If seen, start from the next number that was previously suggested.
        current_counter = generated_filenames_tracker.get(base_name, 0) 

        unique_filename = ""
        full_path_candidate = ""
        
        while True:
            if current_counter == 0:
                unique_filename = proposed_pdf_filename
            else:
                unique_filename = f"{base_name} ({current_counter}){ext}"
            
            full_path_candidate = os.path.join(output_dir, unique_filename)
            
            path_for_check = os.path.abspath(full_path_candidate)

            # Check if this proposed path already exists on disk
            # if os.path.exists(path_for_check):
            #     current_counter += 1
            #     continue # Try next counter
            
            break # Found a unique name

        # Update the tracker with the *next* counter to try for this base_name
        # This ensures that if 'doc.pdf' was used (counter 0), the next time 'doc' is encountered,
        # it will start trying from 'doc (1).pdf' (counter 1).
        generated_filenames_tracker[base_name] = current_counter + 1
        
        return path_for_check

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
        
        # Tracker for unique filenames within this batch
        # Maps base_name (e.g., "document") to the next counter to try (e.g., 1 for "document (1).pdf")
        generated_filenames_tracker = {} 

        try:
            # Ensure the output directory exists.
            if not os.path.isdir(output_dir):
                try:
                    os.makedirs(output_dir)
                    self._log(f"Created output directory: {output_dir}", "blue")
                except Exception as e:
                    self._log(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                    return converted_count, failed_count, total_files

            # Initialize Word Application
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                self._log("Launched a new, isolated Word Application instance.", "blue")
            except Exception as e:
                self._log(f"Error: Could not launch a new, isolated Word Application instance. Please ensure MS Word is installed and not corrupted. Details: {e}", "red")
                return converted_count, failed_count, total_files

            for i, docx_path in enumerate(docx_file_list):
                current_file_name_display = os.path.basename(docx_path)
                self._log(f"[{i+1}/{total_files}] Processing: {current_file_name_display}", "orange")

                doc = None 
                try:
                    if not os.path.exists(docx_path):
                        self._log(f"Skipping '{current_file_name_display}': Source file does not exist.", "red")
                        failed_count += 1
                        continue

                    # Get the initial proposed PDF filename based on the naming rule
                    proposed_pdf_filename = self.get_pdf_filename(docx_path, naming_rule)
                    
                    # Get a unique PDF path, handling duplicates
                    unique_pdf_full_path = self._get_unique_pdf_path(output_dir, proposed_pdf_filename, generated_filenames_tracker)
                    
                    # Extract the filename part for logging
                    unique_pdf_file_name = os.path.basename(unique_pdf_full_path) 

                    # Open Word document
                    doc = word.Documents.Open(
                        os.path.abspath(docx_path), # Use the regular absolute path
                        ReadOnly=True,
                        ConfirmConversions=False,
                        AddToRecentFiles=False
                    )

                    # Save as PDF
                    doc.SaveAs(unique_pdf_full_path, FileFormat=17) 
                    doc.Close(False) 
                    self._log(f"Successfully converted: {current_file_name_display} -> {unique_pdf_file_name}", "green")
                    converted_count += 1

                except pythoncom.com_error as com_e:
                    error_message = f"Conversion of '{current_file_name_display}' failed due to COM error: {com_e}"
                    if hasattr(com_e, 'ex_info') and com_e.ex_info and len(com_e.ex_info) > 1:
                        com_error_description = com_e.ex_info[1]
                        com_error_scode = com_e.ex_info[4]
                        error_message += f"\nDetails: {com_error_description} (HRESULT: {hex(com_error_scode)})"
                        if com_error_scode == -2147024864: # ERROR_SHARING_VIOLATION (0x80070020)
                            error_message += "\nPossible cause: The file is currently in use or locked by another application (e.g., another Word instance). Please close the file and try again."
                        elif com_error_scode == -2147024741: # HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME) (0x8007007B)
                            error_message += "\nPossible cause: The path (source or destination) might be too long or invalid."
                    self._log(error_message, "red")
                    failed_count += 1
                    if doc:
                        try:
                            doc.Close(False)
                        except Exception as close_e:
                            self._log(f"Error closing document after failed COM conversion: {close_e}", "red")

                except Exception as e:
                    error_message = f"Conversion of '{current_file_name_display}' failed: {e}"
                    self._log(error_message, "red")
                    failed_count += 1
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
                    word.Quit() 
                    del word 
                    self._log("Word Application quit and COM object released.", "blue")
                except Exception as e:
                    self._log(f"Error quitting Word application: {e}", "red")

        return converted_count, failed_count, total_files