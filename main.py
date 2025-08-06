# main.py
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk 
import os
import threading

from tkinterdnd2 import DND_FILES, TkinterDnD

from word_to_pdf_converter import WordConverterLogic, BatchConverter

class WordToPdfConverterApp:
    """
    Tkinter GUI application for batch converting Word files to PDF.
    It uses a separate logic class for the conversion process to maintain separation of concerns.
    """
    def __init__(self, master):
        self.master = master
        master.title("Word Batch to PDF Converter")
        master.geometry("700x680")
        master.resizable(False, False)

        master.grid_columnconfigure(1, weight=1)

        self.selected_word_files_data = []
        self.output_pdf_dir = tk.StringVar()

        self.naming_rule_var = tk.StringVar(master)
        self.naming_rules = ["Remove Square Brackets", "Original Name"]
        self.naming_rule_var.set(self.naming_rules[0])
        self.naming_rule_var.trace_add("write", self.on_naming_rule_change)

        self.batch_converter = BatchConverter(log_callback=self.log_status)
        self.converter_logic = WordConverterLogic(log_callback=self.log_status)

        self.summary_window_open = False
        self.master.protocol("WM_DELETE_WINDOW", self.on_main_window_close)

        # --- GUI Control Layout ---

        tk.Label(master, text="Word Files to Convert:").grid(row=0, column=0, padx=10, pady=5, sticky="w")

        tree_frame = tk.Frame(master)
        tree_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        self.word_treeview = ttk.Treeview(tree_frame, columns=("original_word", "converted_pdf"), show="headings", selectmode="extended")

        self.word_treeview.heading("original_word", text="Original Word File")
        self.word_treeview.heading("converted_pdf", text="Converted PDF Name (Preview)")

        self.word_treeview.column("original_word", width=300, anchor="w")
        self.word_treeview.column("converted_pdf", width=300, anchor="w")

        self.word_treeview.grid(row=0, column=0, sticky="nsew")

        self.treeview_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.word_treeview.yview)
        self.treeview_scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.word_treeview.config(yscrollcommand=self.treeview_scrollbar_y.set)

        self.treeview_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.word_treeview.xview)
        self.treeview_scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.word_treeview.config(xscrollcommand=self.treeview_scrollbar_x.set)

        # Bind DND for Treeview
        self.word_treeview.drop_target_register(DND_FILES)
        self.word_treeview.dnd_bind('<<Drop>>', self.handle_treeview_drop)

        # --- File operation buttons with DND frames ---
        # Frame for Add Word Files button to enable DND
        # NEW: Initial border and relief are flat/0
        self.add_files_frame = tk.Frame(master, bd=0, relief="flat") 
        self.add_files_frame.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.add_files_frame.drop_target_register(DND_FILES)
        self.add_files_frame.dnd_bind('<<Drop>>', self.handle_add_files_drop)
        # NEW: Bind DragEnter and DragLeave events
        self.add_files_frame.dnd_bind('<<DragEnter>>', self._on_dnd_enter)
        self.add_files_frame.dnd_bind('<<DragLeave>>', self._on_dnd_leave)

        self.add_files_btn = tk.Button(self.add_files_frame, text="Add Word Files...", command=self.add_word_files)
        self.add_files_btn.pack(padx=5, pady=5)

        self.clear_list_btn = tk.Button(master, text="Clear All", command=self.clear_word_list)
        self.clear_list_btn.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.remove_selected_btn = tk.Button(master, text="Remove Selected", command=self.remove_selected_files)
        self.remove_selected_btn.grid(row=2, column=2, padx=10, pady=5, sticky="w")
        # --- END File operation buttons ---

        # PDF output directory selection
        tk.Label(master, text="Output PDF Directory:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_dir_entry = tk.Entry(master, textvariable=self.output_pdf_dir, width=70)
        self.output_dir_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        # --- Frame for Select Directory button to enable DND ---
        # NEW: Initial border and relief are flat/0
        self.browse_dir_frame = tk.Frame(master, bd=0, relief="flat") 
        self.browse_dir_frame.grid(row=3, column=2, padx=10, pady=5)
        self.browse_dir_frame.drop_target_register(DND_FILES)
        self.browse_dir_frame.dnd_bind('<<Drop>>', self.handle_output_dir_drop)
        # NEW: Bind DragEnter and DragLeave events
        self.browse_dir_frame.dnd_bind('<<DragEnter>>', self._on_dnd_enter)
        self.browse_dir_frame.dnd_bind('<<DragLeave>>', self._on_dnd_leave)

        self.browse_dir_btn = tk.Button(self.browse_dir_frame, text="Select Directory...", command=self.select_output_directory)
        self.browse_dir_btn.pack(padx=5, pady=5)
        # --- END Frame for Select Directory button ---

        # Bind DND for Output Directory Entry (already exists, keep it)
        self.output_dir_entry.drop_target_register(DND_FILES)
        self.output_dir_entry.dnd_bind('<<Drop>>', self.handle_output_dir_drop)

        # PDF Naming Rule selection
        tk.Label(master, text="PDF Naming Rule:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.naming_rule_menu = tk.OptionMenu(master, self.naming_rule_var, *self.naming_rules)
        self.naming_rule_menu.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.naming_rule_menu.config(width=20)

        # --- Modified Section for Centering Buttons ---
        button_frame = tk.Frame(master)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)

        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=0)
        button_frame.grid_columnconfigure(2, weight=0)
        button_frame.grid_columnconfigure(3, weight=1)

        self.convert_btn = tk.Button(button_frame, text="Start Batch Conversion", command=self.start_batch_conversion_thread,
                                     height=2, width=20, bg="lightblue", font=("Arial", 12, "bold"))
        self.convert_btn.grid(row=0, column=1, padx=(0, 10))

        self.stop_btn = tk.Button(button_frame, text="Stop Conversion", command=self.stop_batch_conversion_thread,
                                  height=2, width=15, bg="salmon", font=("Arial", 12, "bold"), state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=2)
        # --- End Modified Section ---

        # Status display area
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
        self.master.after(0, self._update_status_text, message, tag)

    def _update_status_text(self, message, tag):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n", tag)
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)

    def _get_treeview_item_data(self, word_full_path, naming_rule):
        """
        Helper to get the data for a Treeview row (Original Word, Converted PDF).
        """
        word_basename = os.path.basename(word_full_path)
        pdf_filename = self.converter_logic.get_pdf_filename(word_full_path, naming_rule)
        
        return (word_basename, pdf_filename)

    def refresh_treeview_display(self):
        """
        Clears and repopulates the Treeview with current files and naming rule.
        This ensures the preview is always up-to-date.
        """
        for item in self.word_treeview.get_children():
            self.word_treeview.delete(item)
        
        current_naming_rule = self.naming_rule_var.get()
        temp_selected_word_files_data = []
        for item_data in self.selected_word_files_data:
            word_path = item_data['path']
            original_word_name, converted_pdf_name = self._get_treeview_item_data(word_path, current_naming_rule)
            item_id = self.word_treeview.insert("", "end", values=(original_word_name, converted_pdf_name))
            temp_selected_word_files_data.append({'path': word_path, 'treeview_id': item_id})
        self.selected_word_files_data = temp_selected_word_files_data

    def add_word_files(self, file_paths=None):
        """Opens file dialog to select multiple Word files or adds provided paths from DND."""
        if file_paths is None: # If called from button, open dialog
            file_paths = filedialog.askopenfilenames(
                title="Select Word Files",
                filetypes=[
                    ("Word Documents (*.docx)", "*.docx"),
                    ("Word Macro-Enabled Documents (*.docm)", "*.docm"),
                    ("Word 97-2003 Documents (*.doc)", "*.doc"),
                    ("Word Templates (*.dotx;*.dotm;*.dot)", "*.dotx *.dotm *.dot"),
                    ("Rich Text Format (*.rtf)", "*.rtf"),
                    ("OpenDocument Text (*.odt)", "*.odt"),
                    ("All Supported Word Formats", "*.docx *.docm *.doc *.dotx *.dotm *.dot *.rtf *.odt"),
                    ("All Files", "*.*")
                ]
            )
        
        if file_paths:
            added_count = 0
            # Ensure file_paths is iterable and parse it correctly if it's a DND string
            if isinstance(file_paths, str):
                file_paths = self.master.tk.splitlist(file_paths)

            for f_path in file_paths:
                if not os.path.isfile(f_path):
                    self.log_status(f"Dropped item is not a file or does not exist: {f_path}", "orange")
                    continue
                
                valid_extensions = ('.docx', '.docm', '.doc', '.dotx', '.dotm', '.dot', '.rtf', '.odt')
                if not f_path.lower().endswith(valid_extensions):
                    self.log_status(f"Skipping non-Word file: {os.path.basename(f_path)}", "orange")
                    continue

                if not any(data['path'] == f_path for data in self.selected_word_files_data):
                    self.selected_word_files_data.append({'path': f_path, 'treeview_id': None})
                    added_count += 1
            if added_count > 0:
                self.log_status(f"Added {added_count} file(s).", "blue")
                self.refresh_treeview_display()
            else:
                self.log_status("No new files added (might already exist or are not supported Word formats).", "blue")

    def handle_treeview_drop(self, event):
        """Handles files dropped onto the Treeview (file list)."""
        self.log_status(f"Files dropped onto list.", "blue")
        self.add_word_files(event.data)

    def handle_add_files_drop(self, event):
        """Handles files dropped onto the 'Add Word Files' button's frame."""
        self.log_status(f"Files dropped onto 'Add Word Files' button.", "blue")
        self.add_word_files(event.data)
        self._reset_dnd_frame_style(event.widget) # NEW: Reset border after drop

    def handle_output_dir_drop(self, event):
        """Handles directory dropped onto the output directory entry or its button's frame."""
        dropped_paths = self.master.tk.splitlist(event.data)
        if dropped_paths:
            potential_dir = dropped_paths[0]
            if os.path.isdir(potential_dir):
                self.output_pdf_dir.set(potential_dir)
                self.log_status(f"Output directory set by drag-and-drop: {potential_dir}", "blue")
            else:
                self.log_status(f"Dropped item is not a valid directory: {potential_dir}", "orange")
                messagebox.showwarning("Invalid Drop", "Please drop a single directory for the output path.")
        self._reset_dnd_frame_style(event.widget) # NEW: Reset border after drop

    def clear_word_list(self):
        """Clears the Word file list in the GUI and the internal list."""
        self.selected_word_files_data.clear()
        self.word_treeview.delete(*self.word_treeview.get_children())
        self.log_status("File list cleared.", "blue")

    def remove_selected_files(self):
        """Removes selected Word files from the Treeview and internal list."""
        selected_treeview_ids = self.word_treeview.selection()
        if not selected_treeview_ids:
            self.log_status("No files selected to remove.", "orange")
            return

        removed_count = 0
        new_selected_word_files_data = []
        for item_data in self.selected_word_files_data:
            if item_data['treeview_id'] not in selected_treeview_ids:
                new_selected_word_files_data.append(item_data)
            else:
                removed_count += 1
        
        self.selected_word_files_data = new_selected_word_files_data
        
        if removed_count > 0:
            self.log_status(f"Removed {removed_count} selected file(s).", "blue")
            self.refresh_treeview_display()
        else:
            self.log_status("No files were removed.", "blue")

    def select_output_directory(self):
        """Opens directory selection dialog to choose the PDF output directory."""
        dir_path = filedialog.askdirectory(title="Select PDF Output Directory")
        if dir_path:
            self.output_pdf_dir.set(dir_path)
            self.log_status(f"Output directory set to: {dir_path}", "blue")

    def on_naming_rule_change(self, *args):
        """Callback for naming rule dropdown change, refreshes Treeview display."""
        self.refresh_treeview_display()

    def start_batch_conversion_thread(self):
        """
        Prepares for conversion, performs initial validation, and starts the
        conversion process in a separate thread to keep the GUI responsive.
        """
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
        
        if not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir)
                self.log_status(f"Creating output directory: {output_dir}", "blue")
            except Exception as e:
                self.log_status(f"Error: Could not create output directory '{output_dir}': {e}", "red")
                messagebox.showerror("Error", f"Could not create output directory '{output_dir}': {e}")
                return
        
        selected_naming_rule = self.naming_rule_var.get()

        # Disable buttons and update status to indicate conversion is in progress
        self.convert_btn.config(state=tk.DISABLED, text="Converting in progress...", bg="lightgray")
        self.stop_btn.config(state=tk.NORMAL)
        self.add_files_btn.config(state=tk.DISABLED)
        self.clear_list_btn.config(state=tk.DISABLED)
        self.remove_selected_btn.config(state=tk.DISABLED)
        self.browse_dir_btn.config(state=tk.DISABLED)
        self.naming_rule_menu.config(state=tk.DISABLED)
        self.output_dir_entry.config(state=tk.DISABLED)
        self.word_treeview.config(selectmode="none")
        self.log_status("Starting batch conversion...", "blue")

        conversion_thread = threading.Thread(
            target=self._run_conversion_in_thread,
            args=(list(word_paths_for_conversion), output_dir, selected_naming_rule)
        )
        conversion_thread.daemon = True
        conversion_thread.start()

    def stop_batch_conversion_thread(self):
        """
        Calls the stop_conversion method of the BatchConverter to halt the process.
        """
        self.log_status("Attempting to stop conversion...", "orange")
        self.stop_btn.config(state=tk.DISABLED)
        self.batch_converter.stop_conversion()

        # Re-enable GUI elements immediately after signaling stop
        self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue")
        self.add_files_btn.config(state=tk.NORMAL)
        self.clear_list_btn.config(state=tk.NORMAL)
        self.remove_selected_btn.config(state=tk.NORMAL)
        self.browse_dir_btn.config(state=tk.NORMAL)
        self.naming_rule_menu.config(state=tk.NORMAL)
        self.output_dir_entry.config(state=tk.NORMAL)
        self.word_treeview.config(selectmode="extended")
        self.log_status("Conversion stop signal sent. Waiting for workers to finish current tasks.", "orange")


    def _run_conversion_in_thread(self, word_file_list, output_dir, naming_rule):
        """
        Wrapper function to run the conversion logic in a separate thread.
        It calls the BatchConverter and then schedules the final GUI update.
        """
        converted_count, failed_count, total_files = 0, 0, 0
        try:
            final_results, converted_count, failed_count, total_files = self.batch_converter.convert_batch_threaded(
                word_file_list, output_dir, naming_rule
            )
        except Exception as e:
            self.log_status(f"An unexpected error occurred during conversion: {e}", "red")
            final_results = []
        finally:
            self.master.after(0, self._conversion_complete, final_results, converted_count, failed_count, total_files)

    def _conversion_complete(self, final_results, converted_count, failed_count, total_files):
        """
        This method is called on the main Tkinter thread after the conversion thread finishes.
        It re-enables buttons and shows the final summary to the user.
        """
        self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue")
        self.stop_btn.config(state=tk.DISABLED)
        
        self.refresh_treeview_display() 

        final_message = (
            f"Batch conversion complete!\n"
            f"Successfully converted: {converted_count} file(s)\n"
            f"Failed: {failed_count} file(s)\n" 
            f"Total: {total_files} file(s)"
        )
        self.log_status(final_message, "blue")

        if final_results:
            self._show_conversion_summary_window(final_results)

    def _set_main_controls_state(self, state):
        """
        Sets the state of main window controls (buttons, entry, treeview, optionmenu).
        'state' can be tk.NORMAL or tk.DISABLED.
        """
        self.add_files_btn.config(state=state)
        self.clear_list_btn.config(state=state)
        self.remove_selected_btn.config(state=state)
        self.browse_dir_btn.config(state=state)
        self.naming_rule_menu.config(state=state)
        self.output_dir_entry.config(state=state)
        
        if state == tk.NORMAL:
            self.word_treeview.config(selectmode="extended")
        else:
            self.word_treeview.config(selectmode="none")

    # NEW: Method to show border on drag enter
    def _on_dnd_enter(self, event):
        event.widget.config(bd=2, relief="groove")

    # NEW: Method to hide border on drag leave
    def _on_dnd_leave(self, event):
        event.widget.config(bd=0, relief="flat")

    # NEW: Method to reset border after drop (called from drop handlers)
    def _reset_dnd_frame_style(self, widget):
        widget.config(bd=0, relief="flat")

    def _show_conversion_summary_window(self, results):
        """
        Creates a new Toplevel window to display the detailed conversion results
        in a Treeview.
        """
        summary_window = tk.Toplevel(self.master)
        summary_window.title("Conversion Summary")
        summary_window.geometry("800x400")
        summary_window.transient(self.master)

        summary_window.lift()
        summary_window.focus_set()

        self.summary_window_open = True
        summary_window.protocol("WM_DELETE_WINDOW", lambda: self.on_summary_window_close(summary_window))

        self._set_main_controls_state(tk.DISABLED)
        self.convert_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)

        summary_window.grid_rowconfigure(0, weight=1)
        summary_window.grid_columnconfigure(0, weight=1)

        tree_frame = tk.Frame(summary_window)
        tree_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        summary_tree = ttk.Treeview(tree_frame,
                                    columns=("original_file", "converted_pdf", "status", "message"),
                                    show="headings")

        summary_tree.heading("original_file", text="Original File", anchor="w")
        summary_tree.heading("converted_pdf", text="Converted PDF", anchor="w")
        summary_tree.heading("status", text="Status", anchor="center")
        summary_tree.heading("message", text="Message", anchor="w")

        summary_tree.column("original_file", width=200, minwidth=150, stretch=True)
        summary_tree.column("converted_pdf", width=200, minwidth=150, stretch=True)
        summary_tree.column("status", width=80, minwidth=60, stretch=False)
        summary_tree.column("message", width=250, minwidth=100, stretch=True)

        summary_tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=summary_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        summary_tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=summary_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        summary_tree.configure(xscrollcommand=hsb.set)

        for item in results:
            original_file = item.get("original_filename", "N/A")
            converted_pdf = item.get("output_filename", "N/A")
            status = item.get("status", "Unknown")
            message = item.get("message", "")
            renamed = item.get("renamed_due_to_collision", False)

            if renamed:
                tag = "blue"
            elif status == "Success":
                tag = "green"
            elif status == "Failed":
                tag = "red"
            else:
                tag = ""
            summary_tree.insert("", "end", values=(original_file, converted_pdf, status, message), tags=(tag,))
        
        summary_tree.tag_configure("green", foreground="green")
        summary_tree.tag_configure("red", foreground="red")
        summary_tree.tag_configure("blue", foreground="blue")

        close_button = tk.Button(summary_window, text="Close", command=lambda: self.on_summary_window_close(summary_window))
        close_button.grid(row=1, column=0, pady=10)

    def on_summary_window_close(self, summary_window):
        self.summary_window_open = False
        self._set_main_controls_state(tk.NORMAL)
        self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue")
        self.stop_btn.config(state=tk.DISABLED)

        summary_window.destroy()

    def on_main_window_close(self):
        if self.summary_window_open:
            if messagebox.askyesno(
                "Confirm Exit",
                "The conversion summary window is still open. Are you sure you want to exit the application?"
            ):
                for widget in self.master.winfo_children():
                    if isinstance(widget, tk.Toplevel) and widget.title() == "Conversion Summary":
                        widget.destroy()
                        break 
                self._set_main_controls_state(tk.NORMAL)
                self.convert_btn.config(state=tk.NORMAL, text="Start Batch Conversion", bg="lightblue")
                self.stop_btn.config(state=tk.DISABLED)
                self.master.destroy()
        else:
            self.master.destroy()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = WordToPdfConverterApp(root)
    root.mainloop()
