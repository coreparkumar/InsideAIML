# Note: This script now requires the 'thefuzz' library.
# You can install it by running: pip install thefuzz python-Levenshtein
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
import os
import subprocess

# --- Light libraries are imported at the top for fast startup ---

class SearchSetupDialog(simpledialog.Dialog):
    """A custom dialog to get search parameters from the user."""
    def body(self, master):
        self.title("Search Setup")

        # Search Mode
        ttk.Label(master, text="Select Search Mode:").grid(row=0, sticky='w', padx=5, pady=5)
        self.search_mode = tk.StringVar(value="fuzzy")
        
        ttk.Radiobutton(master, text="Fuzzy Match (Similar phrases)", variable=self.search_mode, value="fuzzy", command=self.toggle_threshold).grid(row=1, column=0, sticky='w', padx=20)
        ttk.Radiobutton(master, text="Exact Match (Case-insensitive phrase)", variable=self.search_mode, value="exact", command=self.toggle_threshold).grid(row=2, column=0, sticky='w', padx=20)
        ttk.Radiobutton(master, text="Partial Word Match (Words start with...)", variable=self.search_mode, value="partial", command=self.toggle_threshold).grid(row=3, column=0, sticky='w', padx=20)
        
        # Search Phrase
        ttk.Label(master, text="Search Phrase:").grid(row=4, sticky='w', padx=5, pady=5)
        self.phrase_entry = ttk.Entry(master, width=50)
        self.phrase_entry.grid(row=5, padx=5, sticky='ew')
        self.phrase_entry.focus_set()

        # Similarity Threshold (for Fuzzy Match only)
        self.threshold_label = ttk.Label(master, text="Similarity Threshold (%):")
        self.threshold_label.grid(row=6, sticky='w', padx=5, pady=5)
        self.threshold_spinbox = ttk.Spinbox(master, from_=0, to=100, increment=5, width=10)
        self.threshold_spinbox.set(80)
        self.threshold_spinbox.grid(row=7, padx=5, sticky='w')
        
        self.toggle_threshold()
        return self.phrase_entry

    def toggle_threshold(self):
        """Enable/disable threshold spinbox based on search mode."""
        if self.search_mode.get() == "fuzzy":
            self.threshold_label.config(state='normal')
            self.threshold_spinbox.config(state='normal')
        else:
            self.threshold_label.config(state='disabled')
            self.threshold_spinbox.config(state='disabled')
            
    def apply(self):
        self.result = {
            "mode": self.search_mode.get(),
            "phrase": self.phrase_entry.get().strip(),
            "threshold": int(self.threshold_spinbox.get()) if self.search_mode.get() == "fuzzy" else None
        }

def display_results_and_export(parent, results_data, search_params, base_folder_path):
    """
    Creates a new window to display results, handle interactions, and close the app cleanly.
    """
    import csv

    results_window = tk.Toplevel(parent)
    results_window.title("Search Results")
    results_window.geometry("1100x500")

    # --- FIX: Define a function to properly close the entire application ---
    def on_close():
        """Destroys the main root window, terminating the application."""
        parent.destroy()

    # --- FIX: Bind the window's close button to our custom on_close function ---
    results_window.protocol("WM_DELETE_WINDOW", on_close)

    columns = ("sr_no", "file_name", "match_type", "best_match", "full_path")
    tree = ttk.Treeview(results_window, columns=columns, show="headings")
    
    tree.heading("sr_no", text="Sr. No.")
    tree.heading("file_name", text="File Name")
    tree.heading("match_type", text="Match Info")
    tree.heading("best_match", text="Best Match Found")
    tree.heading("full_path", text="Full Path")

    tree.column("sr_no", width=60, anchor='center')
    tree.column("file_name", width=250)
    tree.column("match_type", width=120, anchor='center')
    tree.column("best_match", width=350)
    tree.column("full_path", width=300)
    
    for i, (display_name, match_data) in enumerate(results_data.items(), 1):
        full_path = os.path.join(base_folder_path, display_name)
        file_name = os.path.basename(display_name)
        match_info = match_data.get('score', 'N/A')
        best_match_text = match_data.get('match', 'N/A')
        tree.insert("", tk.END, values=(i, file_name, match_info, best_match_text, full_path))

    tree.pack(pady=10, padx=10, expand=True, fill='both')

    def open_file_location(event):
        if not tree.selection(): return
        selected_item = tree.selection()[0]
        file_path = tree.item(selected_item, "values")[4]
        if os.path.exists(file_path):
            subprocess.run(['explorer', '/select,', os.path.normpath(file_path)])
    tree.bind("<Double-1>", open_file_location)
    
    def save_to_csv():
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], title="Save Results As CSV")
        if not file_path: return
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['SrNo', 'File Name', 'Search Phrase', 'Match Info', 'Best Match Found', 'File Path'])
                for item in tree.get_children():
                    sr, name, match_info, match_text, path = tree.item(item)['values']
                    writer.writerow([sr, name, search_params['phrase'], match_info, match_text, path])
            messagebox.showinfo("Success", f"Results successfully saved to:\n{file_path}", parent=results_window)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file.\nError: {e}", parent=results_window)
    
    ttk.Button(results_window, text="Export to CSV", command=save_to_csv).pack(pady=10)
    ttk.Label(results_window, text="Double-click a file in the list to open its location.").pack(pady=5)

def search_pdfs_gui():
    root = tk.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(title="Please select a folder with PDF files")
    if not folder_path: return

    setup_dialog = SearchSetupDialog(root)
    search_params = setup_dialog.result
    if not search_params or not search_params.get("phrase"): return
    
    progress_window = tk.Toplevel(root)
    progress_window.title("Working...")
    progress_window.geometry("400x120")
    progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
    
    progress_label = ttk.Label(progress_window, text="Discovering PDF files...")
    progress_label.pack(pady=10)
    
    progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=300, mode='indeterminate')
    progress_bar.pack(pady=10)
    progress_bar.start(10)
    progress_window.update()

    pdf_paths = [os.path.join(dp, f) for dp, _, fn in os.walk(folder_path) for f in fn if f.lower().endswith('.pdf')]
    
    progress_bar.stop()
    progress_bar.config(mode='determinate', maximum=len(pdf_paths), value=0)
    progress_window.title("Searching...")
    
    if not pdf_paths:
        progress_window.destroy()
        messagebox.showinfo("No Files Found", "No PDF files were found.")
        return

    from pypdf import PdfReader
    from thefuzz import fuzz, process
    import re

    results = {}
    search_mode = search_params['mode']
    search_phrase = search_params['phrase']
    
    for i, file_path in enumerate(pdf_paths, 1):
        progress_bar['value'] = i
        progress_label.config(text=f"Scanning ({i}/{len(pdf_paths)}): {os.path.basename(file_path)}")
        progress_window.update_idletasks()

        try:
            reader = PdfReader(file_path)
            full_text = "".join(page.extract_text() or "" for page in reader.pages)
            if not full_text: continue
            
            display_name = os.path.relpath(file_path, folder_path)

            if search_mode == 'fuzzy':
                lines = [line.strip() for line in full_text.splitlines() if line.strip()]
                if lines:
                    best_match, score = process.extractOne(search_phrase, lines, scorer=fuzz.partial_ratio)
                    if score >= search_params['threshold']:
                        results[display_name] = {'match': best_match, 'score': f"{score}%"}
            
            elif search_mode == 'exact':
                for line in full_text.splitlines():
                    if search_phrase.lower() in line.lower():
                        results[display_name] = {'match': line.strip(), 'score': 'Exact Match'}
                        break 
            
            elif search_mode == 'partial':
                search_words = [w for w in search_phrase.split() if len(w) >= 4]
                if not search_words:
                    messagebox.showwarning("Input Error", "Partial match requires search words to be at least 4 letters long.")
                    progress_window.destroy()
                    return
                
                pdf_words = set(re.findall(r'\b\w+\b', full_text.lower()))
                found_matches = []

                for s_word in search_words:
                    for p_word in pdf_words:
                        if p_word.startswith(s_word.lower()):
                            found_matches.append(p_word)
                
                if found_matches:
                    results[display_name] = {'match': ", ".join(sorted(list(set(found_matches)))), 'score': 'Partial Match'}

        except Exception as e:
            print(f"Could not process {file_path}. Error: {e}")

    progress_window.destroy()

    if results:
        display_results_and_export(root, results, search_params, folder_path)
    else:
        messagebox.showinfo("Search Results", "No matches found for the given criteria.")
        root.destroy() # Also destroy root if no results are found

    root.mainloop()

if __name__ == '__main__':
    search_pdfs_gui()

