import os
import win32com.client
import pythoncom
import re
import sys
import threading
import queue
import time # For queue timeout

class WordConverterLogic:
    """
    Handles the core logic for converting WORD files to PDF using MS Word COM automation.
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

    def get_pdf_filename(self, word_path, naming_rule):
        """
        Determines the output PDF filename based on the WORD path and selected naming rule.
        This method is public because the GUI needs to preview the PDF names.
        Note: This method provides the *intended* filename before uniqueness resolution.

        Args:
            word_path (str): The full path to the source WORD file.
            naming_rule (str): The selected naming rule (e.g., "Original Name", "Remove Square Brackets").

        Returns:
            str: The calculated filename for the output PDF.
        """
        base_name = os.path.splitext(os.path.basename(word_path))[0]

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

class ConversionWorker(threading.Thread):
    """
    A worker thread that converts WORD files to PDF using its own Word Application instance.
    """
    def __init__(self, worker_id, task_queue, results_dict, shared_tracker, tracker_lock, output_dir, naming_rule, log_callback, stop_event):
        super().__init__()
        self.worker_id = worker_id
        self.task_queue = task_queue
        self.results_dict = results_dict # Dictionary to store results by original_index
        self.shared_tracker = shared_tracker # Shared dictionary for unique filenames
        self.tracker_lock = tracker_lock # Lock for the shared_tracker
        self.output_dir = output_dir
        self.naming_rule = naming_rule
        self.log_callback = log_callback
        self.stop_event = stop_event # Event to signal stopping
        self.word_app = None
        self.logic = WordConverterLogic(log_callback=self._log) # Each worker gets its own logic instance for logging

    def _log(self, message, tag=None):
        """
        Internal logging method for the worker, prepends worker ID.
        """
        if self.log_callback:
            self.log_callback(f"[Worker {self.worker_id}] {message}", tag)
        else:
            # Fallback to console print if no callback provided
            color_map = {
                "blue": "\033[94m",
                "orange": "\033[93m",
                "green": "\033[92m",
                "red": "\033[91m",
                "reset": "\033[0m"
            }
            colored_message = f"{color_map.get(tag, '')}[Worker {self.worker_id}] {message}{color_map['reset']}"
            print(colored_message)

    def run(self):
        """
        Main execution loop for the worker thread.
        """
        # Initialize COM for this thread
        pythoncom.CoInitialize() 
        self._log("Starting worker.", "blue")
        try:
            while True:
                # Check stop event first. If set and queue is empty, exit.
                if self.stop_event.is_set():
                    self._log("Stop signal received. Exiting worker.", "blue")
                    break

                task = None
                try:
                    # Get a task from the queue with a short timeout.
                    # This timeout allows the worker to check the stop_event periodically.
                    task = self.task_queue.get(timeout=0.1) 
                except queue.Empty:
                    self._log("Queue is empty, no more tasks. Exiting worker.", "blue")
                    break # Exit loop if queue is empty

                original_index = task["original_index"]
                word_path = task["word_path"]
                original_filename = os.path.basename(word_path)
                
                # Initialize result structure for this file
                result = {
                    "original_index": original_index,
                    "original_filename": original_filename,
                    "input_path": word_path,
                    "output_filename": None,
                    "output_path": None,
                    "status": "Failed",
                    "message": "",
                    "renamed_due_to_collision": False
                }


                try:
                    # If stop signal is received *after* getting a task, mark it as failed immediately
                    if self.stop_event.is_set():
                        self._log(f"Stop signal received, marking '{original_filename}' as failed (conversion stopped).", "orange")
                        result["message"] = "Conversion stopped by user."
                        continue # Get next task or exit if queue empty/stop signal still active

                    # Initialize Word Application for this worker if not already done
                    if self.word_app is None:
                        try:
                            self.word_app = win32com.client.DispatchEx("Word.Application")
                            self.word_app.Visible = False
                            self._log("Launched a new, isolated Word Application instance.", "blue")
                        except Exception as e:
                            error_msg = f"Could not launch Word Application instance. Please ensure MS Word is installed and not corrupted. Details: {e}"
                            self._log(error_msg, "red")
                            result["message"] = error_msg
                            # Store result and mark task done before continuing to next task
                            with self.tracker_lock:
                                self.results_dict[original_index] = result
                            # self.task_queue.task_done() # Mark task done even on app launch failure
                            continue # Skip to next task if Word app couldn't be launched
                    
                    # If stop signal is received *after* getting a task, mark it as failed immediately
                    if self.stop_event.is_set():
                        self._log(f"Stop signal received, marking '{original_filename}' as failed (conversion stopped).", "orange")
                        result["message"] = "Conversion stopped by user."
                        continue # Get next task or exit if queue empty/stop signal still active

                    doc = None # Initialize doc object to None
                    try:
                        if not os.path.exists(word_path):
                            error_msg = f"Source file does not exist: '{original_filename}'"
                            self._log(error_msg, "red")
                            result["message"] = error_msg
                            raise FileNotFoundError(error_msg)

                        # Determine proposed PDF filename based on naming rule
                        proposed_pdf_filename = self.logic.get_pdf_filename(word_path, self.naming_rule)
                        
                        # Get a unique PDF path, handling duplicates with a shared tracker and lock
                        final_pdf_full_path, renamed = self._get_unique_pdf_path_thread_safe(
                            self.output_dir, proposed_pdf_filename, self.shared_tracker, self.tracker_lock
                        )
                        
                        final_pdf_filename = os.path.basename(final_pdf_full_path)

                        if len(final_pdf_full_path) > 255:
                            error_msg = (
                                f"Output PDF path is too long ({len(final_pdf_full_path)} characters). "
                                f"Windows path limit is typically 255-260 characters. "
                                f"Please shorten the output directory path or the original filename: '{final_pdf_full_path}'"
                            )
                            self._log(error_msg, "red")
                            result["output_filename"] = final_pdf_filename
                            result["message"] = "Path exceeds 255 chars. Shorten."
                            # This task will be marked as failed in the finally block
                            continue # Skip to next task

                        self._log(f"Processing '{original_filename}' -> '{final_pdf_filename}'", "orange")

                        # Open Word document
                        doc = self.word_app.Documents.Open(
                            os.path.abspath(word_path), # Use the regular absolute path
                            ReadOnly=True,
                            ConfirmConversions=False,
                            AddToRecentFiles=False
                        )

                        # Save as PDF
                        doc.SaveAs(final_pdf_full_path, FileFormat=17) # 17 is wdFormatPDF
                        doc.Close(False) # Close without saving changes to the original WORD
                        doc = None
                        
                        result["status"] = "Success"
                        result["output_filename"] = final_pdf_filename
                        result["output_path"] = final_pdf_full_path
                        result["renamed_due_to_collision"] = renamed
                        result["message"] = "Successfully converted." + (" (Renamed due to collision)" if renamed else "")
                        self._log(f"Successfully converted: '{original_filename}' -> '{final_pdf_filename}'", "green")

                    except pythoncom.com_error as com_e:
                        error_message = f"Conversion of '{original_filename}' failed due to COM error: {com_e}"
                        if hasattr(com_e, 'ex_info') and com_e.ex_info and len(com_e.ex_info) > 1:
                            com_error_description = com_e.ex_info[1]
                            com_error_scode = com_e.ex_info[4]
                            error_message += f"\nDetails: {com_error_description} (HRESULT: {hex(com_error_scode)})"
                            if com_error_scode == -2147024864: # ERROR_SHARING_VIOLATION (0x80070020)
                                error_message += "\nPossible cause: The file is currently in use or locked by another application (e.g., another Word instance). Please close the file and try again."
                            elif com_error_scode == -2147024741: # HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME) (0x8007007B)
                                error_message += "\nPossible cause: The path (source or destination) might be too long or invalid."
                        self._log(error_message, "red")
                        result["message"] = error_message
                        if doc: # Try to close document even if conversion failed
                            try:
                                doc.Close(False)
                            except Exception as close_e:
                                self._log(f"Error closing document after failed COM conversion: {close_e}", "red")

                    except Exception as e:
                        error_message = f"Conversion of '{original_filename}' failed: {e}"
                        self._log(error_message, "red")
                        result["message"] = error_message
                        if doc: # Try to close document even if conversion failed
                            try:
                                doc.Close(False)
                            except Exception as close_e:
                                self._log(f"Error closing document after failed general conversion: {close_e}", "red")

                finally:
                    # Store result in the shared dictionary, protected by the lock
                    with self.tracker_lock:
                        self.results_dict[original_index] = result
                    # Mark the task as done in the queue
                    self.task_queue.task_done()

        finally:
            # Clean up Word application when worker exits its loop
            if self.word_app:
                try:
                    self.word_app.Quit() 
                    del self.word_app 
                    self._log("Word Application quit and COM object released.", "blue")
                except Exception as e:
                    self._log(f"Error quitting Word application: {e}", "red")
            
            # Uninitialize COM for this thread
            pythoncom.CoUninitialize()


    def _get_unique_pdf_path_thread_safe(self, output_dir, proposed_pdf_filename, shared_tracker, tracker_lock):
        """
        Generates a unique PDF path, checking both disk existence and
        names proposed by other threads in the current batch.
        Returns the unique path and a boolean indicating if it was renamed.
        """
        base_name, ext = os.path.splitext(proposed_pdf_filename)
        
        renamed = False
        
        while True:
            current_counter = 0 # Initialize for each attempt in the loop
            with tracker_lock:
                # Get the next counter to try for this base name from the shared tracker.
                # This ensures that if 'doc.pdf' was used, the next time 'doc' is encountered,
                # it will start trying from 'doc (1).pdf'.
                current_counter = shared_tracker.get(base_name, 0)
                
                if current_counter == 0:
                    unique_filename = proposed_pdf_filename
                else:
                    unique_filename = f"{base_name} ({current_counter}){ext}"
                    renamed = True # Mark as renamed if a counter is applied

                full_path_candidate = os.path.join(output_dir, unique_filename)
                path_for_check = os.path.abspath(full_path_candidate)

                # Check if this proposed path already exists on disk
                if os.path.exists(path_for_check):
                    self._log(f"Path '{path_for_check}' already exists on disk. Incrementing counter and retrying.", "orange")
                    # Update tracker to reserve this new counter for this base_name
                    shared_tracker[base_name] = current_counter + 1
                    continue # Try next counter in the next iteration of the while loop

                # If we reach here, the path does not exist on disk.
                # We now "reserve" this name by incrementing the counter in the shared tracker
                # for the *next* time this base_name is requested.
                shared_tracker[base_name] = current_counter + 1
                break # Found a unique name and reserved it

        return path_for_check, renamed


class BatchConverter:
    """
    Orchestrates the multi-threaded batch conversion of WORD files to PDF.
    """
    def __init__(self, log_callback=None):
        self._log_callback = log_callback
        self._stop_event = threading.Event() # Event to signal workers to stop
        self._workers = [] # List to keep track of active worker threads
        self._task_queue = None # Will be initialized in convert_batch_threaded
        self._results_dict = None # Will be initialized in convert_batch_threaded
        self._shared_filename_tracker = None # Will be initialized in convert_batch_threaded
        self._tracker_lock = None # Will be initialized in convert_batch_threaded

    def _log(self, message, tag=None):
        """
        Main logging method for the batch converter.
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

    def _mark_remaining_tasks_as_failed(self):
        """
        Marks any tasks still in the queue as failed and updates results_dict.
        This is called when conversion is explicitly stopped.
        """
        if self._task_queue is not None and self._results_dict is not None and self._tracker_lock is not None:
            while not self._task_queue.empty():
                try:
                    task = self._task_queue.get_nowait() # Get without blocking
                    original_index = task["original_index"]
                    word_path = task["word_path"]
                    original_filename = os.path.basename(word_path)

                    self._log(f"Marking '{original_filename}' as failed (conversion stopped before processing).", "orange")
                    result = {
                        "original_index": original_index,
                        "original_filename": original_filename,
                        "input_path": word_path,
                        "output_filename": None,
                        "output_path": None,
                        "status": "Failed",
                        "message": "Conversion stopped by user before processing.",
                        "renamed_due_to_collision": False
                    }
                    with self._tracker_lock: # Protect access to results_dict
                        self._results_dict[original_index] = result
                    self._task_queue.task_done()
                except queue.Empty:
                    # This should theoretically not happen due to while not empty check, but good practice
                    break 
                except Exception as e:
                    self._log(f"Error marking remaining task as failed: {e}", "red")

    def stop_conversion(self):
        """
        Signals all worker threads to stop and marks any unstarted tasks as failed.
        This method can be called from another thread (e.g., GUI thread).
        """
        if not self._workers:
            self._log("No active conversion to stop.", "orange")
            return

        self._log("Stopping conversion process...", "orange")
        self._stop_event.set() # Set the event to signal all workers to stop

    def convert_batch_threaded(self, word_file_list, output_dir, naming_rule, num_threads=4):
        """
        Performs batch conversion of WORD files to PDF using multiple threads.

        Args:
            word_file_list (list): A list of full paths to WORD files.
            output_dir (str): The directory where converted PDF files will be saved.
            naming_rule (str): The rule to apply for naming the output PDF files.
            num_threads (int): The number of worker threads to use. Defaults to 4.

        Returns:
            tuple: (final_results, converted_count, failed_count, total_files)
                   final_results: A list of dictionaries, each representing the result of a conversion,
                                  ordered by the original input list's index.
        """
        if sys.platform != "win32":
            self._log("This application requires Microsoft Word and pywin32, and therefore only runs on Windows.", "red")
            return [], 0, 0, 0

        if not word_file_list:
            self._log("No WORD files provided for conversion.", "orange")
            return [], 0, 0, 0

        # Ensure the output directory exists.
        if not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir)
                self._log(f"Created output directory: {output_dir}", "blue")
            except Exception as e:
                self._log(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                return [], 0, 0, 0

        # Reset state for a new conversion batch
        self._stop_event.clear() # Ensure stop event is clear for a new run
        self._workers = [] # Clear any old worker references

        self._task_queue = queue.Queue()
        self._results_dict = {}
        self._shared_filename_tracker = {}
        self._tracker_lock = threading.Lock()

        # Pre-process and populate the queue
        self._log(f"Preparing {len(word_file_list)} files for conversion...", "blue")
        for i, word_path in enumerate(word_file_list):
            if not os.path.exists(word_path):
                self._log(f"Skipping '{os.path.basename(word_path)}': Source file does not exist.", "red")
                # Add a failed result immediately for non-existent files
                self._results_dict[i] = {
                    "original_index": i,
                    "original_filename": os.path.basename(word_path),
                    "input_path": word_path,
                    "output_filename": None,
                    "output_path": None,
                    "status": "Failed",
                    "message": "Source file does not exist.",
                    "renamed_due_to_collision": False
                }
                continue

            # Add task to queue
            self._task_queue.put({
                "original_index": i,
                "word_path": word_path,
            })
        
        self._log(f"Queue populated with {self._task_queue.qsize()} tasks.", "blue")

        # Start worker threads
        for i in range(num_threads):
            worker = ConversionWorker(
                worker_id=i + 1,
                task_queue=self._task_queue,
                results_dict=self._results_dict,
                shared_tracker=self._shared_filename_tracker,
                tracker_lock=self._tracker_lock,
                output_dir=output_dir,
                naming_rule=naming_rule,
                log_callback=self._log,
                stop_event=self._stop_event # Pass the stop event to workers
            )
            self._workers.append(worker)
            worker.start()

        # Wait for all worker threads to finish their execution.
        # This will block until all workers naturally complete their tasks or are signaled to stop.
        self._log("Waiting for all worker threads to finish...", "blue")
        for worker in self._workers:
            worker.join()

        # After all worker threads have finished, check the status of the task queue.
        if self._task_queue.empty():
            self._log("All worker threads have finished.", "blue")
        elif self._stop_event.is_set():
            self._mark_remaining_tasks_as_failed()
            self._log("All workers signaled to stop and joined.", "blue")
        else:
            self._mark_remaining_tasks_as_failed()
            self._log("Conversion stopped. All remaining tasks marked as failed.", "orange")

        # Collect and sort results by original_index
        final_results = [self._results_dict[i] for i in sorted(self._results_dict.keys())]
        
        converted_count = sum(1 for r in final_results if r["status"] == "Success")
        failed_count = sum(1 for r in final_results if r["status"] == "Failed")
        total_files = len(final_results)

        self._log(f"Batch conversion complete. Converted: {converted_count}, Failed: {failed_count}, Total: {total_files}", "blue")

        return final_results, converted_count, failed_count, total_files