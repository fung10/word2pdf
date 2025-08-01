# word_to_pdf_converter.py
import os
import win32com.client
import pythoncom
import re
import concurrent.futures
import threading
import queue # Import the queue module

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
        # Using a lock for logging to prevent interleaved messages from multiple threads
        self._log_lock = threading.Lock()

    def _log(self, message, tag=None):
        """
        Internal logging method that uses the provided log_callback.
        If no callback is provided, it prints to console.
        Uses a lock to ensure atomic printing from multiple threads.
        """
        with self._log_lock:
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
            # Corrected regex: matches literal brackets and anything inside them non-greedily
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
        This method is called sequentially during pre-processing, so it's thread-safe.

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
        current_counter = generated_filenames_tracker.get(base_name, 0) 

        full_path_candidate = ""
        
        while True:
            if current_counter == 0:
                unique_filename = proposed_pdf_filename
            else:
                unique_filename = f"{base_name} ({current_counter}){ext}"
            
            full_path_candidate = os.path.abspath(os.path.join(output_dir, unique_filename))
            
            # Check if this proposed path already exists on disk
            if os.path.exists(full_path_candidate):
                current_counter += 1
                continue # Try next counter
            
            break # Found a unique name not existing on disk
            
        # Update the tracker with the *next* counter to try for this base_name
        # This ensures that if 'doc.pdf' was used, the next time 'doc' is encountered,
        # it will start trying from 'doc (1).pdf'.
        generated_filenames_tracker[base_name] = current_counter + 1
        
        return full_path_candidate

    def _process_document_in_word_instance(self, word_app, docx_path, unique_pdf_full_path, file_index, total_files_in_batch):
        """
        Processes a single DOCX file using an *already provided* Word Application instance.
        This function does not handle COM initialization/uninitialization or Word app launching/quitting.

        Args:
            word_app: The already initialized win32com.client.DispatchEx("Word.Application") object.
            docx_path (str): The full path to the source DOCX file.
            unique_pdf_full_path (str): The pre-determined unique full path for the output PDF.
            file_index (int): The 0-based index of the file in the original batch (for logging).
            total_files_in_batch (int): Total number of files in the original batch (for logging).

        Returns:
            bool: True if conversion was successful, False otherwise.
        """
        doc = None # Initialize doc to None
        current_file_name_display = os.path.basename(docx_path)
        unique_pdf_file_name = os.path.basename(unique_pdf_full_path)

        try:
            self._log(f"[{file_index+1}/{total_files_in_batch}] Processing: {current_file_name_display}", "orange")

            # Open Word document using the provided word_app instance
            doc = word_app.Documents.Open(
                os.path.abspath(docx_path),
                ReadOnly=True,
                ConfirmConversions=False,
                AddToRecentFiles=False
            )

            # Save as PDF (FileFormat=17 is wdFormatPDF)
            doc.SaveAs(unique_pdf_full_path, FileFormat=17)
            
            # Close the document immediately after saving
            doc.Close(False) # False means don't save changes to the original DOCX
            doc = None # Crucial: Mark as None after successful close to prevent re-closing in finally

            self._log(f"[{file_index+1}/{total_files_in_batch}] Successfully converted: {current_file_name_display} -> {unique_pdf_file_name}", "green")
            return True

        except pythoncom.com_error as com_e:
            error_message = f"[{file_index+1}/{total_files_in_batch}] Conversion of '{current_file_name_display}' failed due to COM error: {com_e}"
            if hasattr(com_e, 'ex_info') and com_e.ex_info and len(com_e.ex_info) > 1:
                com_error_description = com_e.ex_info[1]
                com_error_scode = com_e.ex_info[4]
                error_message += f"\nDetails: {com_error_description} (HRESULT: {hex(com_error_scode)})"
                if com_error_scode == -2147024864: # ERROR_SHARING_VIOLATION (0x80070020)
                    error_message += "\nPossible cause: The file is currently in use or locked by another application (e.g., another Word instance). Please close the file and try again."
                elif com_error_scode == -2147024741: # HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME) (0x8007007B)
                    error_message += "\nPossible cause: The path (source or destination) might be too long or invalid."
            self._log(error_message, "red")
            return False
        except Exception as e:
            error_message = f"[{file_index+1}/{total_files_in_batch}] Conversion of '{current_file_name_display}' failed: {e}"
            self._log(error_message, "red")
            return False
        finally:
            # If doc is not None here, it means an exception occurred *before* it was closed
            # or before doc was set to None. So, attempt to close it.
            if doc: 
                try:
                    doc.Close(False)
                except Exception as close_e:
                    self._log(f"[{file_index+1}/{total_files_in_batch}] Error closing document in finally block after an error: {close_e}", "red")

    def _worker_thread_conversion(self, task_queue, total_files_in_batch):
        """
        Worker function for a single thread. Initializes one Word instance,
        processes files from the shared task_queue, then quits the Word instance.

        Args:
            task_queue (queue.Queue): The shared queue from which to pull file conversion tasks.
            total_files_in_batch (int): Total number of files in the original batch (for logging).

        Returns:
            tuple: (converted_count_by_worker, failed_count_by_worker)
        """
        word_app = None
        worker_converted_count = 0
        worker_failed_count = 0
        first_task = None

        try:
            pythoncom.CoInitialize() # Initialize COM for this thread

            # Before attempting to launch the Word application, first try to get a task from the queue.
            # This avoids unnecessarily launching Word if a sentinel (None) is received immediately.
            first_task = task_queue.get(block=True)

            if first_task is None: # If the first item retrieved is the sentinel
                task_queue.task_done() # Mark this sentinel task as done
                self._log(f"Worker thread {threading.current_thread().name} received sentinel as first task, exiting without launching Word.", "blue")
                return 0, 0 # Return immediately, do not launch Word application

            # If execution reaches here, first_task is a real task.
            # Now, attempt to launch the Word application.
            try:
                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                self._log(f"Worker thread {threading.current_thread().name} launched a new Word Application instance.", "blue")
            except Exception as e:
                self._log(f"Worker thread {threading.current_thread().name} failed to launch Word Application. Details: {e}", "red")
                # If Word fails to launch, this worker thread cannot process the first_task it retrieved.
                # Re-queue the task so other worker threads can attempt it.
                task_queue.put(first_task)
                self._log(f"Worker thread {threading.current_thread().name} re-queued task due to Word launch failure.", "yellow")
                # It's crucial to call task_done() for the first_task, otherwise task_queue.join() will hang.
                # This signals that this worker has "finished" its responsibility for this specific `get()` operation,
                # even though the task itself is put back for another worker.
                task_queue.task_done()
                # Since Word failed to launch, this worker cannot process any tasks, so it exits.
                return 0, 0 # Return 0 converted and 0 failed, as the task was re-queued

            task = first_task
            while True:                
                if task is None: # Sentinel received, time to exit
                    task_queue.task_done() # Signal that this sentinel task is done
                    break # Exit the worker loop

                docx_path, unique_pdf_full_path, original_idx = task

                try:
                    success = self._process_document_in_word_instance(
                        word_app, docx_path, unique_pdf_full_path, original_idx, total_files_in_batch
                    )
                    if success:
                        worker_converted_count += 1
                    else:
                        worker_failed_count += 1
                finally:
                    task_queue.task_done() # Signal that the task is complete

                # Get a task from the queue. This will block until an item is available.
                # If a None (sentinel) is received, it means it's time to exit.
                task = task_queue.get(block=True) 

        except Exception as e:
            self._log(f"Fatal error in worker thread {threading.current_thread().name}: {e}", "red")
            # If a fatal error occurs, this worker stops.
            # It's important to ensure `task_done()` is called for any task that was `get()`'d, even if processing failed.
            # The structure here ensures that both `first_task` and each `task` in the loop
            # will have `task_done()` called in their respective `finally` blocks.
            # However, if an exception occurs *after* `task_queue.get()` but *before*
            # the `try...finally` block for processing the task (e.g., if `first_task`
            # itself is somehow corrupted and causes an error during unpacking),
            # then `task_done()` might not be called for that specific `get()` operation,
            # potentially leading to `task_queue.join()` hanging.
            # In this modification, we assume `task_queue.get()` itself does not raise
            # an exception that would prevent a subsequent `task_done()`.
        finally:
            if word_app:
                try:
                    word_app.Quit()
                    del word_app
                    self._log(f"Worker thread {threading.current_thread().name} quit its Word Application instance.", "blue")
                except Exception as e:
                    self._log(f"Error quitting Word application in worker thread {threading.current_thread().name}: {e}", "red")
            pythoncom.CoUninitialize() # Uninitialize COM for this thread
        
        return worker_converted_count, worker_failed_count

    def convert_batch(self, docx_file_list, output_dir, naming_rule, max_workers=4):
        """
        Performs batch conversion of DOCX files to PDF using MS Word COM automation,
        processing files concurrently with a fixed number of Word instances and dynamic task distribution.

        Args:
            docx_file_list (list): A list of full paths to DOCX files.
            output_dir (str): The directory where converted PDF files will be saved.
            naming_rule (str): The rule to apply for naming the output PDF files.
            max_workers (int): The maximum number of concurrent Word instances/threads to use.

        Returns:
            tuple: (converted_count, failed_count, total_files)
        """
        converted_count = 0
        failed_count = 0
        initial_total_files = len(docx_file_list) # Store original total for final report
        
        if not docx_file_list:
            self._log("No DOCX files provided for conversion.", "orange")
            return 0, 0, 0

        # Ensure the output directory exists.
        if not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir)
                self._log(f"Created output directory: {output_dir}", "blue")
            except Exception as e:
                self._log(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                return converted_count, failed_count, initial_total_files

        # --- Pre-process to determine all unique output filenames sequentially ---
        # This ensures that naming conflicts (e.g., two input files resulting in "doc.pdf")
        # are resolved before concurrent processing, preventing race conditions on filenames.
        files_to_process_info = [] # Store (docx_path, unique_pdf_full_path, original_idx)
        generated_filenames_tracker = {}

        self._log("Pre-processing files to determine unique output paths...", "blue")
        for i, docx_path in enumerate(docx_file_list):
            if not os.path.exists(docx_path):
                self._log(f"Skipping '{os.path.basename(docx_path)}': Source file does not exist. It will be counted as failed.", "red")
                failed_count += 1 # Count as failed if source file doesn't exist
                continue

            proposed_pdf_filename = self.get_pdf_filename(docx_path, naming_rule)
            unique_pdf_full_path = self._get_unique_pdf_path(output_dir, proposed_pdf_filename, generated_filenames_tracker)
            # Store info for task: (docx_path, unique_pdf_full_path, original_idx)
            # total_files_in_batch is passed to worker function, not per task.
            files_to_process_info.append((docx_path, unique_pdf_full_path, i)) 

        total_files_for_processing = len(files_to_process_info)
        if total_files_for_processing == 0 and failed_count == initial_total_files:
            self._log("No valid DOCX files found for conversion after pre-processing (all were skipped or non-existent).", "orange")
            return converted_count, failed_count, initial_total_files
        elif total_files_for_processing == 0:
            self._log("No valid DOCX files found for conversion after pre-processing.", "orange")
            return converted_count, failed_count, initial_total_files

        self._log(f"Starting concurrent conversion of {total_files_for_processing} files with up to {max_workers} workers...", "blue")

        task_queue = queue.Queue()
        for file_info in files_to_process_info:
            task_queue.put(file_info)

        # Use ThreadPoolExecutor to run worker functions
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="WordConverterWorker") as executor:
            futures = []
            # Submit max_workers worker functions. Each worker will pull from the queue.
            for _ in range(max_workers):
                # Pass the queue and total_files_in_batch (for logging context) to each worker
                futures.append(executor.submit(self._worker_thread_conversion, task_queue, initial_total_files))

            # Wait for all tasks in the queue to be processed
            # This will block until task_done() has been called for every item put() into the queue.
            task_queue.join() 

            # After all real tasks are done, put a 'None' sentinel for each worker to signal exit.
            for _ in range(max_workers):
                task_queue.put(None) 

            # Wait for all worker futures to complete (i.e., workers exit their loops after processing sentinels)
            for future in concurrent.futures.as_completed(futures):
                try:
                    worker_converted, worker_failed = future.result()
                    converted_count += worker_converted
                    failed_count += worker_failed
                except Exception as exc:
                    self._log(f"A worker thread generated an unexpected exception: {exc}", "red")
                    pass # The worker's return value already includes its failures.
        
        self._log(f"Batch conversion complete. Converted: {converted_count}, Failed: {failed_count}, Total Files Attempted (including pre-processing skips): {initial_total_files}", "blue")
        # Return counts relative to the *original* list of files provided by the user
        return converted_count, failed_count, initial_total_files