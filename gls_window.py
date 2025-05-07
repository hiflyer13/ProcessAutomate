import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


class GLSWindow:
    def __init__(self, window, main_window):
        self.window = window
        self.main_window = main_window

        # Configure the window
        self.window.configure(bg="#f0f0f0")

        # Add title
        self.title_label = tk.Label(
            window,
            text="GLS",
            font=("Arial", 24, "bold"),
            bg="#f0f0f0"
        )
        self.title_label.pack(pady=30)

        # Files label
        self.files_label = tk.Label(
            window,
            text="Files",
            font=("Arial", 14),
            bg="#f0f0f0"
        )
        self.files_label.pack(anchor=tk.W, padx=50, pady=10)

        # Frame for file entry and scrollbar
        self.files_frame = tk.Frame(window, bg="#f0f0f0")
        self.files_frame.pack(fill=tk.BOTH, expand=True, padx=50)

        # File entry box with scrollbar
        self.file_listbox = tk.Listbox(self.files_frame, width=50, height=5)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

        self.scrollbar = tk.Scrollbar(self.files_frame, orient=tk.VERTICAL)
        self.scrollbar.config(command=self.file_listbox.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox.config(yscrollcommand=self.scrollbar.set)

        # Button to browse files
        self.browse_button = tk.Button(
            window,
            text="Browse Files",
            font=("Arial", 12),
            bg="#4a7abc",
            fg="white",
            width=20,
            height=2,
            command=self.browse_files
        )
        self.browse_button.pack(pady=10)

        # Optional label
        self.optional_label = tk.Label(
            window,
            text="Optional",
            font=("Arial", 14),
            bg="#f0f0f0"
        )
        self.optional_label.pack(anchor=tk.W, padx=50, pady=10)

        # Frame for optional entry boxes
        self.optional_frame = tk.Frame(window, bg="#f0f0f0")
        self.optional_frame.pack(fill=tk.BOTH, expand=True, padx=50)

        # First optional entry (text)
        self.optional_entry1 = tk.Entry(self.optional_frame, width=30)
        self.optional_entry1.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame for second optional entry with negative sign
        self.optional2_frame = tk.Frame(self.optional_frame, bg="#f0f0f0")
        self.optional2_frame.pack(side=tk.LEFT, padx=5, pady=5)

        # Negative sign label
        self.negative_label = tk.Label(
            self.optional2_frame,
            text="-",
            font=("Arial", 12),
            bg="#f0f0f0"
        )
        self.negative_label.pack(side=tk.LEFT)

        # Second optional entry (positive integer only)
        self.optional_entry2_var = tk.StringVar()
        self.optional_entry2_var.trace("w", self.validate_positive_integer)
        self.optional_entry2 = tk.Entry(
            self.optional2_frame,
            width=29,  # Slightly smaller to account for the negative sign
            textvariable=self.optional_entry2_var
        )
        self.optional_entry2.pack(side=tk.LEFT)

        # Run button
        self.run_button = tk.Button(
            window,
            text="Run",
            font=("Arial", 12),
            bg="#4a7abc",
            fg="white",
            width=20,
            height=2,
            command=self.run_function
        )
        self.run_button.pack(pady=10)

        # Back button
        self.back_button = tk.Button(
            window,
            text="Back to Main Menu",
            font=("Arial", 12),
            bg="#4a7abc",
            fg="white",
            width=20,
            height=2,
            command=self.go_back
        )
        self.back_button.pack(pady=10)

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.go_back)

    def validate_positive_integer(self, *args):
        """Validate that the input is a positive integer or empty"""
        value = self.optional_entry2_var.get()
        if value == "":
            return  # Allow empty field

        try:
            # Ensure it's a positive integer
            int_value = int(value)
            if int_value < 0:
                # Remove negative sign if user tries to enter one
                self.optional_entry2_var.set(value.replace('-', ''))
        except ValueError:
            # If not an integer, remove the last character
            self.optional_entry2_var.set(value[:-1])

    def browse_files(self):
        # Open file dialog to select multiple files
        file_paths = filedialog.askopenfilenames(title="Select Files")
        for file_path in file_paths:
            self.file_listbox.insert(tk.END, file_path)

    def run_function(self):
        # Get all files from the listbox instead of just selected ones
        all_files = [self.file_listbox.get(idx) for idx in range(self.file_listbox.size())]

        if not all_files:
            messagebox.showwarning("No Files", "Please browse and add files to process.")
            return

        optional_1 = self.optional_entry1.get()
        optional_2 = self.optional_entry2.get()

        # Convert optional_2 to negative integer if it's not empty
        if optional_2:
            try:
                optional_2 = -int(optional_2)  # Make it negative
            except ValueError:
                messagebox.showerror("Invalid Input", "Optional field 2 must be a number.")
                return

        processed_files = []
        errors = []

        for file_path in all_files:
            try:
                # Load and process each file
                print(f"Processing file: {file_path}")
                df = pd.read_excel(file_path)
                # Remove the first 7 rows
                df = df.drop(index=range(0, 7))
                # Remove the last row:
                df = df.drop(df.index[-1])
                # Reset the row index
                df.reset_index(drop=True, inplace=True)
                # Reset column index
                df.columns = range(df.shape[1])
                # Keeping columns 2 and 4
                df = df.iloc[:, [2, 4]]
                # Converting the grand total to integer
                df[4] = df[4].astype(int)

                # Append optional entries if provided
                if optional_1 or optional_2:
                    optional_data = [optional_1, optional_2]
                    df.loc[df.shape[0]] = optional_data

                # Construct the output file path
                dir_name = os.path.dirname(file_path)
                base_name = os.path.basename(file_path).replace('.xlsx', '')
                output_file_path = os.path.join(dir_name, f"processed_{base_name}.xlsx")

                print(f"Saving to: {output_file_path}")

                # Save the DataFrame to a new Excel file with "processed_" prefix
                # Using pandas ExcelWriter to avoid the warning
                with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')

                processed_files.append(output_file_path)
                print(f"Successfully saved: {output_file_path}")

            except Exception as e:
                error_msg = f"Error processing {file_path}: {str(e)}"
                errors.append(error_msg)
                print(error_msg)

        # Show appropriate message based on results
        if errors:
            error_text = "\n".join(errors)
            messagebox.showerror("Errors Occurred", f"The following errors occurred:\n{error_text}")
        elif processed_files:
            processed_text = "\n".join(processed_files)
            messagebox.showinfo("Success", f"All files processed successfully!\nFiles saved:\n{processed_text}")
        else:
            messagebox.showwarning("No Files Processed", "No files were processed.")

    def go_back(self):
        # Show the main window again
        self.main_window.deiconify()
        # Close this window
        self.window.destroy()