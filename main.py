# main.py
import tkinter as tk # Keep standard tkinter import
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk 
import os
import threading

# Import TkinterDnD2 for DND functionality
# We need to import TkinterDnD2.Tk() for the root window
# and DND_FILES for drop target registration.
from tkinterdnd2 import DND_FILES, TkinterDnD # Import specifically for DND features

# Import the logic module
# Assuming word_to_pdf_converter.py exists and contains WordConverterLogic and BatchConverter
from word_to_pdf_converter import WordConverterLogic, BatchConverter

class WordToPdfConverterApp:
    """
    Tkinter GUI application for batch converting Word files to PDF.
    It uses a separate logic class for the conversion process to maintain separation of concerns.
    """
    def __init__(self, master):
        # master should be an instance of TkinterDnD2.Tk()
        self.master = master
        master.title("Word Batch to PDF Converter")
        master.geometry("700x680")
        master.resizable(False, False)

        # Configure column weights for grid layout
        master.grid_columnconfigure(1, weight=1)

        self.selected_word_files_data = []
        self.output_pdf_dir = tk.StringVar() # Use tk.StringVar

        self.naming_rule_var = tk.StringVar(master) # Use tk.StringVar
        self.naming_rules = ["Remove Square Brackets", "Original Name"]
        self.naming_rule_var.set(self.naming_rules[0])
        self.naming_rule_var.trace_add("write", self.on_naming_rule_change)

        self.batch_converter = BatchConverter(log_callback=self.log_status)
        self.converter_logic = WordConverterLogic(log_callback=self.log_status)

        # --- GUI Control Layout ---

        tk.Label(master, text="Word Files to Convert:").grid(row=0, column=0, padx=10, pady=5, sticky="w") # Use tk.Label

        tree_frame = tk.Frame(master) # Use tk.Frame
        tree_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        self.word_treeview = ttk.Treeview(tree_frame, columns=("original_word", "converted_pdf"), show="headings", selectmode="extended") # Changed from docx_treeview

        self.word_treeview.heading("original_word", text="Original Word File") # Changed heading
        self.word_treeview.heading("converted_pdf", text="Converted PDF Name (Preview)")

        self.word_treeview.column("original_word", width=300, anchor="w") # Changed column
        self.word_treeview.column("converted_pdf", width=300, anchor="w")

        self.word_treeview.grid(row=0, column=0, sticky="nsew")

        self.treeview_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.word_treeview.yview) # Changed treeview
        self.treeview_scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.word_treeview.config(yscrollcommand=self.treeview_scrollbar_y.set) # Changed treeview

        self.treeview_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.word_treeview.xview) # Changed treeview
        self.treeview_scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.word_treeview.config(xscrollcommand=self.treeview_scrollbar_x.set) # Changed treeview

        # Bind DND for Treeview
        self.word_treeview.drop_target_register(DND_FILES) # Changed treeview, Use DND_FILES
        self.word_treeview.dnd_bind('<<Drop>>', self.handle_treeview_drop) # Changed treeview, Bind the drop event

        # File operation buttons
        self.add_files_btn = tk.Button(master, text="Add Word Files...", command=self.add_word_files)
        self.add_files_btn.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.clear_list_btn = tk.Button(master, text="Clear All", command=self.clear_word_list)
        self.clear_list_btn.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.remove_selected_btn = tk.Button(master, text="Remove Selected", command=self.remove_selected_files) # Use tk.Button
        self.remove_selected_btn.grid(row=2, column=2, padx=10, pady=5, sticky="w")

        # PDF output directory selection
        tk.Label(master, text="Output PDF Directory:").grid(row=3, column=0, padx=10, pady=5, sticky="w") # Use tk.Label
        self.output_dir_entry = tk.Entry(master, textvariable=self.output_pdf_dir, width=70) # Use tk.Entry
        self.output_dir_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.browse_dir_btn = tk.Button(master, text="Select Directory...", command=self.select_output_directory) # Use tk.Button
        self.browse_dir_btn.grid(row=3, column=2, padx=10, pady=5)

        # Bind DND for Output Directory Entry
        self.output_dir_entry.drop_target_register(DND_FILES) # Use DND_FILES
        self.output_dir_entry.dnd_bind('<<Drop>>', self.handle_output_dir_drop) # Bind the drop event

        # PDF Naming Rule selection
        tk.Label(master, text="PDF Naming Rule:").grid(row=4, column=0, padx=10, pady=5, sticky="w") # Use tk.Label
        self.naming_rule_menu = tk.OptionMenu(master, self.naming_rule_var, *self.naming_rules) # Use tk.OptionMenu
        self.naming_rule_menu.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.naming_rule_menu.config(width=20)

        # --- Modified Section for Centering Buttons ---
        button_frame = tk.Frame(master)
        # Place the frame spanning all columns (0, 1, 2)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)

        # Configure columns within the button_frame to center its contents
        button_frame.grid_columnconfigure(0, weight=1) # Left spacer column
        button_frame.grid_columnconfigure(1, weight=0) # Column for convert_btn
        button_frame.grid_columnconfigure(2, weight=0) # Column for stop_btn
        button_frame.grid_columnconfigure(3, weight=1) # Right spacer column

        self.convert_btn = tk.Button(button_frame, text="Start Batch Conversion", command=self.start_batch_conversion_thread,
                                     height=2, width=20, bg="lightblue", font=("Arial", 12, "bold"))
        # Place convert_btn in column 1 of button_frame, with right padding for the gap
        self.convert_btn.grid(row=0, column=1, padx=(0, 10))

        self.stop_btn = tk.Button(button_frame, text="Stop Conversion", command=self.stop_batch_conversion_thread,
                                  height=2, width=15, bg="salmon", font=("Arial", 12, "bold"), state=tk.DISABLED)
        # Place stop_btn in column 2 of button_frame
        self.stop_btn.grid(row=0, column=2)
        # --- End Modified Section ---

        # Status display area
        tk.Label(master, text="Conversion Log/Status:").grid(row=6, column=0, padx=10, pady=5, sticky="w") # Use tk.Label
        self.status_text = scrolledtext.ScrolledText(master, width=80, height=8, state=tk.DISABLED, wrap=tk.WORD) # Use tk.DISABLED, tk.WORD
        self.status_text.grid(row=7, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

        # Configure tags for colored logging
        self.status_text.tag_config("green", foreground="green")
        self.status_text.tag_config("red", foreground="red")
        self.status_text.tag_config("blue", foreground="blue")
        self.status_text.tag_config("orange", foreground="orange")

        # Initial display update
        self.refresh_treeview_display()

    def log_status(self, message, tag=None):
        """
        Adds log messages to the status text box.
        This method is thread-safe as it uses master.after() to update the GUI
        from any thread that calls it.
        """
        self.master.after(0, self._update_status_text, message, tag)

    def _update_status_text(self, message, tag):
        """Actual GUI update for status text, called from the main thread."""
        self.status_text.config(state=tk.NORMAL) # Use tk.NORMAL
        self.status_text.insert(tk.END, message + "\n", tag) # Use tk.END
        self.status_text.see(tk.END) # Scroll to the latest message
        self.status_text.config(state=tk.DISABLED) # Use tk.DISABLED

    def _get_treeview_item_data(self, word_full_path, naming_rule):
        """
        Helper to get the data for a Treeview row (Original Word, Converted PDF).
        """
        word_basename = os.path.basename(word_full_path)
        # Use the converter_logic's method to get the PDF filename preview
        pdf_filename = self.converter_logic.get_pdf_filename(word_full_path, naming_rule)
        
        return (word_basename, pdf_filename)

    def refresh_treeview_display(self):
        """
        Clears and repopulates the Treeview with current files and naming rule.
        This ensures the preview is always up-to-date.
        """
        # Clear existing items
        for item in self.word_treeview.get_children(): # Changed treeview
            self.word_treeview.delete(item)
        
        current_naming_rule = self.naming_rule_var.get()
        # Create a temporary list to rebuild selected_word_files_data with updated treeview_id
        temp_selected_word_files_data = []
        for item_data in self.selected_word_files_data:
            word_path = item_data['path']
            original_word_name, converted_pdf_name = self._get_treeview_item_data(word_path, current_naming_rule)
            # Insert item and store its Treeview ID back into our data structure
            item_id = self.word_treeview.insert("", "end", values=(original_word_name, converted_pdf_name)) # Changed treeview
            temp_selected_word_files_data.append({'path': word_path, 'treeview_id': item_id})
        self.selected_word_files_data = temp_selected_word_files_data

    # Modified add_word_files to accept an optional file_paths argument for DND
    def add_word_files(self, file_paths=None):
        """Opens file dialog to select multiple Word files or adds provided paths from DND."""
        if file_paths is None: # If called from button, open dialog
            file_paths = filedialog.askopenfilenames(
                title="Select Word Files", # Changed dialog title
                filetypes=[
                    ("Word Documents (*.docx)", "*.docx"),
                    ("Word Macro-Enabled Documents (*.docm)", "*.docm"),
                    ("Word 97-2003 Documents (*.doc)", "*.doc"),
                    ("Word Templates (*.dotx;*.dotm;*.dot)", "*.dotx *.dotm *.dot"),
                    ("Rich Text Format (*.rtf)", "*.rtf"),
                    ("OpenDocument Text (*.odt)", "*.odt"),
                    ("All Supported Word Formats", "*.docx *.docm *.doc *.dotx *.dotm *.dot *.rtf *.odt"), # Consolidated
                    ("All Files", "*.*")
                ]
            )
        
        if file_paths:
            added_count = 0
            # Ensure file_paths is iterable and parse it correctly if it's a DND string
            if isinstance(file_paths, str):
                # TkinterDnD2's event.data can return a space-separated string of paths
                # Use master.tk.splitlist to handle paths with spaces correctly
                file_paths = self.master.tk.splitlist(file_paths)

            for f_path in file_paths:
                # Basic check for file existence and common Word file extensions
                if not os.path.isfile(f_path):
                    self.log_status(f"Dropped item is not a file or does not exist: {f_path}", "orange")
                    continue
                
                # Check if it's a common Word document extension
                # This list should ideally match the filetypes in askopenfilenames
                valid_extensions = ('.docx', '.docm', '.doc', '.dotx', '.dotm', '.dot', '.rtf', '.odt')
                if not f_path.lower().endswith(valid_extensions):
                    self.log_status(f"Skipping non-Word file: {os.path.basename(f_path)}", "orange")
                    continue

                # Check if the file path already exists in our data list
                if not any(data['path'] == f_path for data in self.selected_word_files_data):
                    # For new items, treeview_id is None initially. It will be set by refresh_treeview_display.
                    self.selected_word_files_data.append({'path': f_path, 'treeview_id': None})
                    added_count += 1
            if added_count > 0:
                self.log_status(f"Added {added_count} file(s).", "blue")
                self.refresh_treeview_display() # Refresh the entire Treeview display
            else:
                self.log_status("No new files added (might already exist or are not supported Word formats).", "blue")

    def handle_treeview_drop(self, event):
        """Handles files dropped onto the Treeview (file list)."""
        # event.data contains the paths of the dropped files/folders
        self.log_status(f"Files dropped onto list: {event.data}", "blue")
        self.add_word_files(event.data)

    def handle_output_dir_drop(self, event):
        """Handles directory dropped onto the output directory entry."""
        dropped_paths = self.master.tk.splitlist(event.data)
        if dropped_paths:
            # We only expect one directory to be dropped for the output path
            # If multiple items are dropped, take the first one and check if it's a directory
            potential_dir = dropped_paths[0]
            if os.path.isdir(potential_dir):
                self.output_pdf_dir.set(potential_dir)
                self.log_status(f"Output directory set by drag-and-drop: {potential_dir}", "blue")
            else:
                self.log_status(f"Dropped item is not a valid directory: {potential_dir}", "orange")
                messagebox.showwarning("Invalid Drop", "Please drop a single directory for the output path.")

    def clear_word_list(self):
        """Clears the Word file list in the GUI and the internal list."""
        self.selected_word_files_data.clear()
        self.word_treeview.delete(*self.word_treeview.get_children()) # Changed treeview
        self.log_status("File list cleared.", "blue")

    def remove_selected_files(self):
        """Removes selected Word files from the Treeview and internal list."""
        selected_treeview_ids = self.word_treeview.selection() # Changed treeview
        if not selected_treeview_ids:
            self.log_status("No files selected to remove.", "orange")
            return

        removed_count = 0
        # Create a new list for files that are NOT being removed
        new_selected_word_files_data = []
        for item_data in self.selected_word_files_data:
            if item_data['treeview_id'] not in selected_treeview_ids:
                new_selected_word_files_data.append(item_data)
            else:
                removed_count += 1
        
        self.selected_word_files_data = new_selected_word_files_data
        
        if removed_count > 0:
            self.log_status(f"Removed {removed_count} selected file(s).", "blue")
            self.refresh_treeview_display() # Refresh the entire Treeview display
        else:
            self.log_status("No files were removed.", "blue")

    def select_output_directory(self):
        """Opens directory selection dialog to choose the PDF output directory."""
        dir_path = filedialog.askdirectory(title="Select PDF Output Directory")
        if dir_path:
            self.output_pdf_dir.set(dir_path)
            self.log_status(f"Output directory set to: {dir_path}", "blue")

    def on_naming_rule_change(self, *args): # Callback for trace_add receives args (var_name, index, mode)
        """Callback for naming rule dropdown change, refreshes Treeview display."""
        self.refresh_treeview_display()

    def start_batch_conversion_thread(self):
        """
        Prepares for conversion, performs initial validation, and starts the
        conversion process in a separate thread to keep the GUI responsive.
        """
        # Get the actual list of full paths to pass to the converter logic
        word_paths_for_conversion = [item_data['path'] for item_data in self.selected_word_files_data]

        if not word_paths_for_conversion:
            self.log_status("Error: Please add Word files first.", "red")
            messagebox.showerror("Error", "Please add Word files for conversion.")
            return

        output_dir = self.output_pdf_dir.get()
        if not output_dir:
            self.log_status("Error: Please select an output directory.", "red")
            messagebox.showerror("Error", "Please select an output directory to save the converted PDF files.")
            return
        
        # Check if output directory exists or can be created (perform this check on main thread)
        # The logic module also checks this, but doing it here provides quicker feedback to the user.
        if not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir) # Attempt to create directory
                self.log_status(f"Creating output directory: {output_dir}", "blue")
            except Exception as e:
                self.log_status(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                messagebox.showerror("Error", f"Could not create output directory '{output_dir}': {e}")
                return
        
        selected_naming_rule = self.naming_rule_var.get() # Get the selected naming rule

        # Disable buttons and update status to indicate conversion is in progress
        self.convert_btn.config(state=tk.DISABLED, text="Converting in progress...", bg="lightgray") # Use tk.DISABLED
        self.stop_btn.config(state=tk.NORMAL) # Enable stop button
        self.add_files_btn.config(state=tk.DISABLED)
        self.clear_list_btn.config(state=tk.DISABLED)
        self.remove_selected_btn.config(state=tk.DISABLED)
        self.browse_dir_btn.config(state=tk.DISABLED)
        self.naming_rule_menu.config(state=tk.DISABLED)
        self.word_treeview.config(selectmode="none") # Disable selection during conversion # Changed treeview
        self.log_status("Starting batch conversion...", "blue")

        # Create and start a new thread to run the conversion logic.
        conversion_thread = threading.Thread(
            target=self._run_conversion_in_thread,
            args=(list(word_paths_for_conversion), output_dir, selected_naming_rule) # Pass a copy
        )
        conversion_thread.daemon = True # Allow the program to exit even if thread is running
        conversion_thread.start()

    def stop_batch_conversion_thread(self):
        """
        Calls the stop_conversion method of the BatchConverter to halt the process.
        """
        self.log_status("Attempting to stop conversion...", "orange")
        self.stop_btn.config(state=tk.DISABLED) # Disable stop button immediately to prevent multiple clicks
        self.batch_converter.stop_conversion() # Signal the batch converter to stop

        # Re-enable GUI elements immediately after signaling stop
        # The _conversion_complete will also run, but this provides quicker feedback
        self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue")
        self.add_files_btn.config(state=tk.NORMAL)
        self.clear_list_btn.config(state=tk.NORMAL)
        self.remove_selected_btn.config(state=tk.NORMAL)
        self.browse_dir_btn.config(state=tk.NORMAL)
        self.naming_rule_menu.config(state=tk.NORMAL)
        self.word_treeview.config(selectmode="extended") # Re-enable selection # Changed treeview
        self.log_status("Conversion stop signal sent. Waiting for workers to finish current tasks.", "orange")


    def _run_conversion_in_thread(self, word_file_list, output_dir, naming_rule):
        """
        Wrapper function to run the conversion logic in a separate thread.
        It calls the BatchConverter and then schedules the final GUI update.
        """
        converted_count, failed_count, total_files = 0, 0, 0
        try:
            # Call the conversion logic from the separate thread, passing the naming rule
            final_results, converted_count, failed_count, total_files = self.batch_converter.convert_batch_threaded(
                word_file_list, output_dir, naming_rule
            )
        except Exception as e:
            self.log_status(f"An unexpected error occurred during conversion: {e}", "red")
            final_results = [] # Ensure final_results is defined even on error
        finally:
            # Schedule the final UI update to run on the main Tkinter thread
            self.master.after(0, self._conversion_complete, final_results, converted_count, failed_count, total_files)

    def _conversion_complete(self, final_results, converted_count, failed_count, total_files):
        """
        This method is called on the main Tkinter thread after the conversion thread finishes.
        It re-enables buttons and shows the final summary to the user.
        """
        # Re-enable GUI elements
        self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue") # Use tk.NORMAL
        self.stop_btn.config(state=tk.DISABLED) # Disable stop button
        self.add_files_btn.config(state=tk.NORMAL)
        self.clear_list_btn.config(state=tk.NORMAL)
        self.remove_selected_btn.config(state=tk.NORMAL)
        self.browse_dir_btn.config(state=tk.NORMAL)
        self.naming_rule_menu.config(state=tk.NORMAL)
        self.word_treeview.config(selectmode="extended") # Re-enable selection # Changed treeview

        # Refresh Treeview display to reflect any changes (e.g., if files were processed)
        # Note: The current logic doesn't remove successfully converted files from the list,
        # but a more advanced version could update their status in the Treeview.
        self.refresh_treeview_display() 

        # Display final summary
        final_message = (
            f"Batch conversion complete!\n"
            f"Successfully converted: {converted_count} file(s)\n"
            f"Failed: {failed_count} file(s)\n" 
            f"Total: {total_files} file(s)"
        )
        self.log_status(final_message, "blue")
        # messagebox.showinfo("Batch Conversion Complete", final_message)

        # --- NEW: Show detailed results in a new window ---
        if final_results: # Only show if there are results to display
            self._show_conversion_summary_window(final_results)
        # --- END NEW ---

    def _show_conversion_summary_window(self, results):
        """
        Creates a new Toplevel window to display the detailed conversion results
        in a Treeview.
        """
        summary_window = tk.Toplevel(self.master)
        summary_window.title("Conversion Summary")
        summary_window.geometry("800x400")
        summary_window.transient(self.master) # Make it appear on top of the main window
        summary_window.grab_set() # Make it modal (user must interact with it before main window)

        # Configure grid for the summary window
        summary_window.grid_rowconfigure(0, weight=1)
        summary_window.grid_columnconfigure(0, weight=1)

        # Frame for the Treeview and scrollbars
        tree_frame = tk.Frame(summary_window)
        tree_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Create the Treeview
        summary_tree = ttk.Treeview(tree_frame,
                                    columns=("original_file", "converted_pdf", "status", "message"),
                                    show="headings")

        # Define headings
        summary_tree.heading("original_file", text="Original File", anchor="w")
        summary_tree.heading("converted_pdf", text="Converted PDF", anchor="w")
        summary_tree.heading("status", text="Status", anchor="center")
        summary_tree.heading("message", text="Message", anchor="w")

        # Define column widths (adjust as needed)
        summary_tree.column("original_file", width=200, minwidth=150, stretch=True)
        summary_tree.column("converted_pdf", width=200, minwidth=150, stretch=True)
        summary_tree.column("status", width=80, minwidth=60, stretch=False)
        summary_tree.column("message", width=250, minwidth=100, stretch=True)

        summary_tree.grid(row=0, column=0, sticky="nsew")

        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=summary_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        summary_tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=summary_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        summary_tree.configure(xscrollcommand=hsb.set)

        # Populate the Treeview with results
        for item in results:
            original_file = item.get("original_filename", "N/A")
            converted_pdf = item.get("output_filename", "N/A")
            status = item.get("status", "Unknown")
            message = item.get("message", "")
            renamed = item.get("renamed_due_to_collision", False)

            # You can add tags for coloring based on status if desired
            if renamed:
                tag = "blue"
            elif status == "Success":
                tag = "green"
            elif status == "Failed":
                tag = "red"
            else:
                tag = ""
            summary_tree.insert("", "end", values=(original_file, converted_pdf, status, message), tags=(tag,))
        
        # Apply tags for coloring
        summary_tree.tag_configure("green", foreground="green")
        summary_tree.tag_configure("red", foreground="red")
        summary_tree.tag_configure("blue", foreground="blue")

        # Add a close button
        close_button = tk.Button(summary_window, text="Close", command=summary_window.destroy)
        close_button.grid(row=1, column=0, pady=10)

        # Set focus to the summary window and wait until it's closed
        summary_window.wait_window()


if __name__ == "__main__":
    root = TkinterDnD.Tk() # IMPORTANT: Use TkinterDnD.Tk() for the root window
    app = WordToPdfConverterApp(root) # Changed from DocxToPdfConverterApp
    root.mainloop()