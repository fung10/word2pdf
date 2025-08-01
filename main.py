# main.py
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk # Import ttk module for Treeview
import os
import threading

# Import the logic module
from word_to_pdf_converter import DocxConverterLogic

class DocxToPdfConverterApp:
    """
    Tkinter GUI application for batch converting DOCX files to PDF.
    It uses a separate logic class for the conversion process to maintain separation of concerns.
    """
    def __init__(self, master):
        self.master = master
        master.title("DOCX Batch to PDF Converter")
        master.geometry("700x680") # Adjusted initial window size for Treeview
        master.resizable(False, False)

        # Configure column weights for grid layout
        master.grid_columnconfigure(1, weight=1)

        # Tkinter variables to store file paths and directories
        # self.selected_docx_files_data will now store dictionaries with 'path' and 'treeview_id'
        self.selected_docx_files_data = [] # List to store dicts: {'path': full_path, 'treeview_id': item_id}
        self.output_pdf_dir = tk.StringVar()
        
        # Variable for naming rule selection
        self.naming_rule_var = tk.StringVar(master)
        # These naming rules must match the strings expected by DocxConverterLogic.get_pdf_filename
        self.naming_rules = ["Original Name", "Remove Square Brackets"] 
        self.naming_rule_var.set(self.naming_rules[0]) # Set default value
        
        # Trace changes to naming_rule_var to update Treeview immediately
        self.naming_rule_var.trace_add("write", self.on_naming_rule_change)

        # Initialize the conversion logic, passing the GUI's log_status method as a callback
        self.converter_logic = DocxConverterLogic(log_callback=self.log_status)

        # --- GUI Control Layout ---

        # Treeview for file list
        tk.Label(master, text="DOCX Files to Convert:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Frame to hold Treeview and Scrollbar
        tree_frame = tk.Frame(master)
        tree_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1) # Allow Treeview to expand
        tree_frame.grid_rowconfigure(0, weight=1) # Allow Treeview to expand

        # Define Treeview
        self.docx_treeview = ttk.Treeview(tree_frame, columns=("original_docx", "converted_pdf"), show="headings", selectmode="extended")
        
        # Define column headings
        self.docx_treeview.heading("original_docx", text="Original DOCX File")
        self.docx_treeview.heading("converted_pdf", text="Converted PDF Name (Preview)")
        
        # Define column widths and anchoring
        self.docx_treeview.column("original_docx", width=300, anchor="w")
        self.docx_treeview.column("converted_pdf", width=300, anchor="w")
        
        self.docx_treeview.grid(row=0, column=0, sticky="nsew")

        # Scrollbar for the Treeview
        self.treeview_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.docx_treeview.yview)
        self.treeview_scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.docx_treeview.config(yscrollcommand=self.treeview_scrollbar_y.set)

        self.treeview_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.docx_treeview.xview)
        self.treeview_scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.docx_treeview.config(xscrollcommand=self.treeview_scrollbar_x.set)

        # File operation buttons
        self.add_files_btn = tk.Button(master, text="Add DOCX Files...", command=self.add_docx_files)
        self.add_files_btn.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        
        self.clear_list_btn = tk.Button(master, text="Clear All", command=self.clear_docx_list)
        self.clear_list_btn.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        self.remove_selected_btn = tk.Button(master, text="Remove Selected", command=self.remove_selected_files)
        self.remove_selected_btn.grid(row=2, column=2, padx=10, pady=5, sticky="w")

        # PDF output directory selection
        tk.Label(master, text="Output PDF Directory:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_dir_entry = tk.Entry(master, textvariable=self.output_pdf_dir, width=70)
        self.output_dir_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.browse_dir_btn = tk.Button(master, text="Select Directory...", command=self.select_output_directory)
        self.browse_dir_btn.grid(row=3, column=2, padx=10, pady=5)

        # PDF Naming Rule selection
        tk.Label(master, text="PDF Naming Rule:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.naming_rule_menu = tk.OptionMenu(master, self.naming_rule_var, *self.naming_rules)
        self.naming_rule_menu.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.naming_rule_menu.config(width=20) 

        # Convert button
        self.convert_btn = tk.Button(master, text="Start Batch Conversion", command=self.start_batch_conversion_thread,
                                     height=2, width=20, bg="lightblue", font=("Arial", 12, "bold"))
        self.convert_btn.grid(row=5, column=0, columnspan=3, pady=20)

        # Status display area (using ScrolledText for multi-line logs)
        tk.Label(master, text="Conversion Log/Status:").grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.status_text = scrolledtext.ScrolledText(master, width=80, height=8, state=tk.DISABLED, wrap=tk.WORD)
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
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n", tag)
        self.status_text.see(tk.END) # Scroll to the latest message
        self.status_text.config(state=tk.DISABLED)

    def _get_treeview_item_data(self, docx_full_path, naming_rule):
        """
        Helper to get the data for a Treeview row (Original DOCX, Converted PDF).
        """
        docx_basename = os.path.basename(docx_full_path)
        # Use the converter_logic's method to get the PDF filename preview
        pdf_filename = self.converter_logic.get_pdf_filename(docx_full_path, naming_rule)
        
        return (docx_basename, pdf_filename)

    def refresh_treeview_display(self):
        """
        Clears and repopulates the Treeview with current files and naming rule.
        This ensures the preview is always up-to-date.
        """
        # Clear existing items
        for item in self.docx_treeview.get_children():
            self.docx_treeview.delete(item)
        
        current_naming_rule = self.naming_rule_var.get()
        # Create a temporary list to rebuild selected_docx_files_data with updated treeview_id
        temp_selected_docx_files_data = []
        for item_data in self.selected_docx_files_data:
            docx_path = item_data['path']
            original_docx_name, converted_pdf_name = self._get_treeview_item_data(docx_path, current_naming_rule)
            # Insert item and store its Treeview ID back into our data structure
            item_id = self.docx_treeview.insert("", "end", values=(original_docx_name, converted_pdf_name))
            temp_selected_docx_files_data.append({'path': docx_path, 'treeview_id': item_id})
        self.selected_docx_files_data = temp_selected_docx_files_data # Update the main list

    def add_docx_files(self):
        """Opens file dialog to select multiple DOCX files and adds them to the list."""
        file_paths = filedialog.askopenfilenames(
            title="Select DOCX Files",
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
            for f_path in file_paths:
                # Check if the file path already exists in our data list
                if not any(data['path'] == f_path for data in self.selected_docx_files_data):
                    # For new items, treeview_id is None initially. It will be set by refresh_treeview_display.
                    self.selected_docx_files_data.append({'path': f_path, 'treeview_id': None}) 
                    added_count += 1
            if added_count > 0:
                self.log_status(f"Added {added_count} file(s).", "blue")
                self.refresh_treeview_display() # Refresh the entire Treeview display
            else:
                self.log_status("No new files added (might already exist).", "blue")

    def clear_docx_list(self):
        """Clears the DOCX file list in the GUI and the internal list."""
        self.selected_docx_files_data.clear()
        self.docx_treeview.delete(*self.docx_treeview.get_children()) # Clear all items from Treeview
        self.log_status("File list cleared.", "blue")

    def remove_selected_files(self):
        """Removes selected DOCX files from the Treeview and internal list."""
        selected_treeview_ids = self.docx_treeview.selection()
        if not selected_treeview_ids:
            self.log_status("No files selected to remove.", "orange")
            return

        removed_count = 0
        # Create a new list for files that are NOT being removed
        new_selected_docx_files_data = []
        for item_data in self.selected_docx_files_data:
            if item_data['treeview_id'] not in selected_treeview_ids:
                new_selected_docx_files_data.append(item_data)
            else:
                removed_count += 1
        
        self.selected_docx_files_data = new_selected_docx_files_data
        
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
        docx_paths_for_conversion = [item_data['path'] for item_data in self.selected_docx_files_data]

        if not docx_paths_for_conversion:
            self.log_status("Error: Please add DOCX files first.", "red")
            messagebox.showerror("Error", "Please add DOCX files for conversion.")
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
        self.convert_btn.config(state=tk.DISABLED, text="Converting in progress...", bg="lightgray")
        self.add_files_btn.config(state=tk.DISABLED)
        self.clear_list_btn.config(state=tk.DISABLED)
        self.remove_selected_btn.config(state=tk.DISABLED)
        self.browse_dir_btn.config(state=tk.DISABLED)
        self.naming_rule_menu.config(state=tk.DISABLED)
        self.docx_treeview.config(selectmode="none") # Disable selection during conversion
        self.log_status("Starting batch conversion...", "blue")

        # Create and start a new thread to run the conversion logic.
        conversion_thread = threading.Thread(
            target=self._run_conversion_in_thread,
            args=(list(docx_paths_for_conversion), output_dir, selected_naming_rule) # Pass a copy
        )
        conversion_thread.daemon = True # Allow the program to exit even if thread is running
        conversion_thread.start()

    def _run_conversion_in_thread(self, docx_file_list, output_dir, naming_rule):
        """
        Wrapper function to run the conversion logic in a separate thread.
        It calls the DocxConverterLogic and then schedules the final GUI update.
        """
        converted_count, failed_count, total_files = 0, 0, 0
        try:
            # Call the conversion logic from the separate thread, passing the naming rule
            converted_count, failed_count, total_files = self.converter_logic.convert_batch(
                docx_file_list, output_dir, naming_rule
            )
        except Exception as e:
            self.log_status(f"An unexpected error occurred during conversion: {e}", "red")
        finally:
            # Schedule the final UI update to run on the main Tkinter thread
            self.master.after(0, self._conversion_complete, converted_count, failed_count, total_files)

    def _conversion_complete(self, converted_count, failed_count, total_files):
        """
        This method is called on the main Tkinter thread after the conversion thread finishes.
        It re-enables buttons and shows the final summary to the user.
        """
        # Re-enable GUI elements
        self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue")
        self.add_files_btn.config(state=tk.NORMAL)
        self.clear_list_btn.config(state=tk.NORMAL)
        self.remove_selected_btn.config(state=tk.NORMAL)
        self.browse_dir_btn.config(state=tk.NORMAL)
        self.naming_rule_menu.config(state=tk.NORMAL)
        self.docx_treeview.config(selectmode="extended") # Re-enable selection

        # Refresh Treeview display to reflect any changes (e.g., if files were processed)
        # Note: The current logic doesn't remove successfully converted files from the list,
        # but a more advanced version could update their status in the Treeview.
        self.refresh_treeview_display() 

        # Display final summary
        final_message = (
            f"Batch conversion complete!\n"
            f"Successfully converted: {converted_count} file(s)\n"
            f"Failed: {failed_count} file(s)"
        )
        self.log_status(final_message, "blue")
        messagebox.showinfo("Batch Conversion Complete", final_message)


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfConverterApp(root)
    root.mainloop()