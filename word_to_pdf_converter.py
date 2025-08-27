import os
import win32com.client
import pythoncom
import re
import sys
import threading
import queue
import time

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
            cleaned_base_name = re.sub(r'\[.*?\]', '', base_name)
            cleaned_base_name = re.sub(r'\s+', ' ', cleaned_base_name).strip()
            if not cleaned_base_name:
                cleaned_base_name = "Untitled_Document"
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
        self.results_dict = results_dict
        self.shared_tracker = shared_tracker
        self.tracker_lock = tracker_lock
        self.output_dir = output_dir
        self.naming_rule = naming_rule
        self.log_callback = log_callback
        self.stop_event = stop_event
        self.word_app = None
        self.logic = WordConverterLogic(log_callback=self._log)

    def _log(self, message, tag=None):
        """
        Internal logging method for the worker, prepends worker ID.
        """
        if self.log_callback:
            self.log_callback(f"[Worker {self.worker_id}] {message}", tag)
        else:
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
        pythoncom.CoInitialize() 
        self._log("Starting worker.", "blue")
        try:
            while True:
                if self.stop_event.is_set():
                    self._log("Stop signal received. Exiting worker.", "blue")
                    break

                task = None
                try:
                    task = self.task_queue.get(timeout=0.1) 
                except queue.Empty:
                    self._log("Queue is empty, no more tasks. Exiting worker.", "blue")
                    break

                original_index = task["original_index"]
                word_path = task["word_path"]
                original_filename = os.path.basename(word_path)
                
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
                    if self.stop_event.is_set():
                        self._log(f"Stop signal received, marking '{original_filename}' as failed (conversion stopped).", "orange")
                        result["message"] = "Conversion stopped by user."
                        continue

                    if self.word_app is None:
                        try:
                            self.word_app = win32com.client.DispatchEx("Word.Application")
                            self.word_app.Visible = False
                            self._log("Launched a new, isolated Word Application instance.", "blue")
                        except Exception as e:
                            error_msg = f"Could not launch Word Application instance. Please ensure MS Word is installed and not corrupted. Details: {e}"
                            self._log(error_msg, "red")
                            result["message"] = error_msg
                            with self.tracker_lock:
                                self.results_dict[original_index] = result
                            continue
                    
                    if self.stop_event.is_set():
                        self._log(f"Stop signal received, marking '{original_filename}' as failed (conversion stopped).", "orange")
                        result["message"] = "Conversion stopped by user."
                        continue

                    doc = None
                    try:
                        if not os.path.exists(word_path):
                            error_msg = f"Source file does not exist: '{original_filename}'"
                            self._log(error_msg, "red")
                            result["message"] = error_msg
                            raise FileNotFoundError(error_msg)

                        proposed_pdf_filename = self.logic.get_pdf_filename(word_path, self.naming_rule)
                        
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
                            continue

                        self._log(f"Processing '{original_filename}' -> '{final_pdf_filename}'", "orange")

                        doc = self.word_app.Documents.Open(
                            os.path.abspath(word_path),
                            ReadOnly=True,
                            ConfirmConversions=False,
                            AddToRecentFiles=False
                        )

                        doc.SaveAs(final_pdf_full_path, FileFormat=17)
                        doc.Close(False)
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
                            if com_error_scode == -2147024864:
                                error_message += "\nPossible cause: The file is currently in use or locked by another application (e.g., another Word instance). Please close the file and try again."
                            elif com_error_scode == -2147024741:
                                error_message += "\nPossible cause: The path (source or destination) might be too long or invalid."
                        self._log(error_message, "red")
                        result["message"] = error_message
                        if doc:
                            try:
                                doc.Close(False)
                            except Exception as close_e:
                                self._log(f"Error closing document after failed COM conversion: {close_e}", "red")

                    except Exception as e:
                        error_message = f"Conversion of '{original_filename}' failed: {e}"
                        self._log(error_message, "red")
                        result["message"] = error_message
                        if doc:
                            try:
                                doc.Close(False)
                            except Exception as close_e:
                                self._log(f"Error closing document after failed general conversion: {close_e}", "red")

                finally:
                    with self.tracker_lock:
                        self.results_dict[original_index] = result
                    self.task_queue.task_done()

        finally:
            if self.word_app:
                try:
                    self.word_app.Quit() 
                    del self.word_app 
                    self._log("Word Application quit and COM object released.", "blue")
                except Exception as e:
                    self._log(f"Error quitting Word application: {e}", "red")
            
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
            current_counter = 0
            with tracker_lock:
                current_counter = shared_tracker.get(base_name, 0)
                
                if current_counter == 0:
                    unique_filename = proposed_pdf_filename
                else:
                    unique_filename = f"{base_name} ({current_counter}){ext}"
                    renamed = True

                full_path_candidate = os.path.join(output_dir, unique_filename)
                path_for_check = os.path.abspath(full_path_candidate)

                if os.path.exists(path_for_check):
                    self._log(f"Path '{path_for_check}' already exists on disk. Incrementing counter and retrying.", "orange")
                    shared_tracker[base_name] = current_counter + 1
                    continue

                shared_tracker[base_name] = current_counter + 1
                break

        return path_for_check, renamed


class BatchConverter:
    """
    Orchestrates the multi-threaded batch conversion of WORD files to PDF.
    """
    def __init__(self, log_callback=None):
        self._log_callback = log_callback
        self._stop_event = threading.Event()
        self._workers = []
        self._task_queue = None
        self._results_dict = None
        self._shared_filename_tracker = None
        self._tracker_lock = None

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
                    task = self._task_queue.get_nowait()
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
                    with self._tracker_lock:
                        self._results_dict[original_index] = result
                    self._task_queue.task_done()
                except queue.Empty:
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
        self._stop_event.set()

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

        if not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir)
                self._log(f"Created output directory: {output_dir}", "blue")
            except Exception as e:
                self._log(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                return [], 0, 0, 0

        self._stop_event.clear()
        self._workers = []

        self._task_queue = queue.Queue()
        self._results_dict = {}
        self._shared_filename_tracker = {}
        self._tracker_lock = threading.Lock()

        self._log(f"Preparing {len(word_file_list)} files for conversion...", "blue")
        for i, word_path in enumerate(word_file_list):
            if not os.path.exists(word_path):
                self._log(f"Skipping '{os.path.basename(word_path)}': Source file does not exist.", "red")
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

            self._task_queue.put({
                "original_index": i,
                "word_path": word_path,
            })
        
        self._log(f"Queue populated with {self._task_queue.qsize()} tasks.", "blue")

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
                stop_event=self._stop_event
            )
            self._workers.append(worker)
            worker.start()

        self._log("Waiting for all worker threads to finish...", "blue")
        for worker in self._workers:
            worker.join()

        if self._task_queue.empty():
            self._log("All worker threads have finished.", "blue")
        elif self._stop_event.is_set():
            self._mark_remaining_tasks_as_failed()
            self._log("All workers signaled to stop and joined.", "blue")
        else:
            self._mark_remaining_tasks_as_failed()
            self._log("Conversion stopped. All remaining tasks marked as failed.", "orange")

        final_results = [self._results_dict[i] for i in sorted(self._results_dict.keys())]
        
        converted_count = sum(1 for r in final_results if r["status"] == "Success")
        failed_count = sum(1 for r in final_results if r["status"] == "Failed")
        total_files = len(final_results)

        self._log(f"Batch conversion complete. Converted: {converted_count}, Failed: {failed_count}, Total: {total_files}", "blue")

        return final_results, converted_count, failed_count, total_files
