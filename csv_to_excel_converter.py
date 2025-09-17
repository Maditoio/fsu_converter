"""
CSV to Excel Converter with Lock ID Conversion
Author: Mumba Mukendi and Github Copilot
Description: A simple GUI application that converts CSV files to Excel format
             and replaces lock values with lock id in hexadecimal format.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from pathlib import Path
import threading
import subprocess
import platform


class CSVToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("FSU CSV to Excel Converter")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_files = []
        self.output_directory = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="FSU HEXER TOOL", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Description
        desc_text = ("Select CSV files with format: lock,fsu,start,end\n"
                    "The app will convert to Excel and replace 'lock' with 'lock id' in hex format")
        desc_label = ttk.Label(main_frame, text=desc_text, 
                              font=("Arial", 10), foreground="gray")
        desc_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(0, weight=1)
        
        # File list
        self.file_listbox = tk.Listbox(file_frame, height=6, selectmode=tk.EXTENDED)
        self.file_listbox.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File selection buttons
        select_btn = ttk.Button(file_frame, text="Select CSV Files", 
                               command=self.select_files)
        select_btn.grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        
        clear_btn = ttk.Button(file_frame, text="Clear Selection", 
                              command=self.clear_files)
        clear_btn.grid(row=1, column=1, sticky=tk.W)
        
        # Output directory section
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Output Directory:").grid(row=0, column=0, sticky=tk.W)
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_directory, state="readonly")
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        
        browse_btn = ttk.Button(output_frame, text="Browse", 
                               command=self.select_output_directory)
        browse_btn.grid(row=0, column=2, sticky=tk.W)
        
        # Convert section
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.convert_btn = ttk.Button(convert_frame, text="Convert Files", 
                                     command=self.convert_files, 
                                     style="Accent.TButton")
        self.convert_btn.pack()
        
        # Progress section
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to convert files")
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Set default output directory to user's Documents
        default_output = str(Path.home() / "Documents" / "FSU_Converted_Files")
        self.output_directory.set(default_output)
    
    def select_files(self):
        """Open file dialog to select CSV files"""
        files = filedialog.askopenfilenames(
            title="Select CSV Files",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if files:
            self.selected_files = list(files)
            self.update_file_list()
    
    def clear_files(self):
        """Clear the selected files list"""
        self.selected_files = []
        self.update_file_list()
    
    def update_file_list(self):
        """Update the file listbox with selected files"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.selected_files:
            filename = os.path.basename(file_path)
            self.file_listbox.insert(tk.END, filename)
    
    def select_output_directory(self):
        """Select output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_directory.set(directory)
    
    def _open_directory(self, directory_path):
        """Open directory in the system's default file manager (cross-platform)"""
        try:
            system = platform.system()
            if system == "Darwin":    # macOS
                subprocess.run(["open", directory_path])
            elif system == "Windows": # Windows
                os.startfile(directory_path)
            elif system == "Linux":   # Linux
                subprocess.run(["xdg-open", directory_path])
            else:
                print(f"Cannot open directory on {system}. Directory: {directory_path}")
        except Exception as e:
            messagebox.showwarning("Cannot Open Directory", 
                                 f"Could not open directory: {directory_path}\nError: {e}")
    
    def convert_files(self):
        """Convert selected CSV files to Excel with lock id column (hex format)"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select at least one CSV file.")
            return
        
        if not self.output_directory.get():
            messagebox.showwarning("No Output Directory", "Please select an output directory.")
            return
        
        # Disable convert button and start progress
        self.convert_btn.config(state="disabled")
        self.progress_bar.start()
        self.progress_label.config(text="Converting files...")
        
        # Run conversion in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._convert_files_thread)
        thread.daemon = True
        thread.start()
    
    def _convert_files_thread(self):
        """Thread function for file conversion"""
        try:
            output_dir = Path(self.output_directory.get())
            output_dir.mkdir(parents=True, exist_ok=True)
            
            converted_files = []
            failed_files = []
            
            for file_path in self.selected_files:
                try:
                    self._update_progress(f"Converting {os.path.basename(file_path)}...")
                    
                    # Read CSV file
                    df = pd.read_csv(file_path)
                    
                    # Validate required columns
                    required_columns = ['lock', 'fsu', 'start', 'end']
                    if not all(col in df.columns for col in required_columns):
                        raise ValueError(f"Missing required columns. Expected: {required_columns}")
                    
                    # Add lock id column (hex without 0x prefix)
                    df['lock id'] = df['lock'].apply(lambda x: hex(int(x))[2:] if pd.notna(x) else '')
                    
                    # Remove the original lock column and reorder
                    df = df.drop('lock', axis=1)
                    columns = ['lock id', 'fsu', 'start', 'end']
                    df = df[columns]
                    
                    # Generate output filename
                    input_name = Path(file_path).stem
                    output_file = output_dir / f"{input_name}_converted.xlsx"
                    
                    # Save to Excel
                    df.to_excel(output_file, index=False)
                    converted_files.append(output_file)
                    
                except Exception as e:
                    failed_files.append((file_path, str(e)))
            
            # Update UI in main thread
            self.root.after(0, self._conversion_complete, converted_files, failed_files)
            
        except Exception as e:
            self.root.after(0, self._conversion_error, str(e))
    
    def _update_progress(self, message):
        """Update progress label from thread"""
        self.root.after(0, lambda: self.progress_label.config(text=message))
    
    def _conversion_complete(self, converted_files, failed_files):
        """Handle conversion completion"""
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")
        
        # Show results
        if converted_files and not failed_files:
            message = f"Successfully converted {len(converted_files)} file(s)!\n\n"
            message += f"Files saved to:\n{self.output_directory.get()}"
            self.progress_label.config(text=f"Conversion complete! {len(converted_files)} file(s) converted.")
            messagebox.showinfo("Conversion Complete", message)
            
            # Ask if user wants to open the output directory
            if messagebox.askyesno("Open Directory", "Would you like to open the output directory?"):
                self._open_directory(self.output_directory.get())
                
        elif converted_files and failed_files:
            message = f"Converted {len(converted_files)} file(s) successfully.\n"
            message += f"Failed to convert {len(failed_files)} file(s):\n\n"
            for file_path, error in failed_files:
                message += f"• {os.path.basename(file_path)}: {error}\n"
            self.progress_label.config(text=f"Partial success: {len(converted_files)} converted, {len(failed_files)} failed.")
            messagebox.showwarning("Partial Success", message)
            
        else:
            message = "Failed to convert any files:\n\n"
            for file_path, error in failed_files:
                message += f"• {os.path.basename(file_path)}: {error}\n"
            self.progress_label.config(text="Conversion failed.")
            messagebox.showerror("Conversion Failed", message)
    
    def _conversion_error(self, error_message):
        """Handle conversion error"""
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")
        self.progress_label.config(text="Conversion failed.")
        messagebox.showerror("Error", f"An error occurred during conversion:\n{error_message}")


def main():
    """Main function to run the application"""
    root = tk.Tk()
    
    # Set application icon (optional)
    try:
        root.iconbitmap(default="icon.ico")
    except:
        pass
    
    # Create and run the application
    app = CSVToExcelConverter(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main()