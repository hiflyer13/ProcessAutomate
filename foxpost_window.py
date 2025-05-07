import os
import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QListWidget, QMessageBox, \
    QFileDialog


class FoxpostWindow:
    def __init__(self, window, main_window):
        self.window = window
        self.main_window = main_window

        # Set up central widget and layout
        self.central_widget = QWidget()
        self.window.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Add title
        self.title_label = QLabel("Foxpost")
        self.title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.layout.addWidget(self.title_label)

        # Files label
        self.files_label = QLabel("Files")
        self.files_label.setStyleSheet("font-size: 14px;")
        self.layout.addWidget(self.files_label)

        # File listbox
        self.file_listbox = QListWidget()
        self.layout.addWidget(self.file_listbox)

        # Button to browse files
        self.browse_button = QPushButton("Browse Files")
        self.browse_button.setStyleSheet("background-color: #4a7abc; color: white; font-size: 12px; padding: 10px;")
        self.browse_button.clicked.connect(self.browse_files)
        self.layout.addWidget(self.browse_button)

        # Run button
        self.run_button = QPushButton("Run")
        self.run_button.setStyleSheet("background-color: #4a7abc; color: white; font-size: 12px; padding: 10px;")
        self.run_button.clicked.connect(self.run_function)
        self.layout.addWidget(self.run_button)

        # Back button
        self.back_button = QPushButton("Back to Main Menu")
        self.back_button.setStyleSheet("background-color: #4a7abc; color: white; font-size: 12px; padding: 10px;")
        self.back_button.clicked.connect(self.go_back)
        self.layout.addWidget(self.back_button)

        # Connect close event
        self.window.closeEvent = self.handle_close_event

    def browse_files(self):
        # Open file dialog to select multiple files
        file_paths, _ = QFileDialog.getOpenFileNames(self.window, "Select Files")
        for file_path in file_paths:
            self.file_listbox.addItem(file_path)

    def run_function(self):
        # Get all files from the listbox
        all_files = [self.file_listbox.item(i).text() for i in range(self.file_listbox.count())]

        if not all_files:
            QMessageBox.warning(self.window, "No Files", "Please browse and add files to process.")
            return

        processed_files = []
        errors = []

        for file_path in all_files:
            try:
                # Load and process each file
                print(f"Processing file: {file_path}")

                # Step 1: Process the 'utánvétek' sheet
                df_utanvetek = pd.read_excel(file_path,
                                             sheet_name='utánvétek',  # The sheet name is fixed
                                             skiprows=10,
                                             header=None)  # Skip the first 10 rows to start from row 11
                df_utanvetek_filtered = df_utanvetek[[4, 7]]

                # Step 2: Process the 'összesítés' sheet
                # First, read the entire sheet without skipping rows
                raw_data = pd.read_excel(file_path, sheet_name='összesítés', header=None)  # Sheet name is fixed

                # Find the row containing "ÖSSZESÍTÉS"
                target_row = -1
                for i, row in enumerate(raw_data.values):
                    # Convert row to string and check if "ÖSSZESÍTÉS" is in any cell
                    row_as_str = [str(cell) for cell in row]
                    if any("ÖSSZESÍTÉS" in cell for cell in row_as_str):
                        target_row = i
                        break

                # If found, create a DataFrame starting from the row after "ÖSSZESÍTÉS"
                if target_row != -1:
                    # Read the Excel file again, but now skip rows up to and including the target row
                    df_new = pd.read_excel(file_path,
                                           sheet_name='összesítés',
                                           skiprows=target_row + 1,
                                           header=None)

                    filtered_rows = df_new[df_new[0].astype(str).str.contains("PARTNER", na=False)]
                    # Change the value to 1
                    filtered_rows[0] = 1
                    # Convert the number to integer
                    filtered_rows[1] = filtered_rows[1].astype(int)
                    # Keeping the relevant columns only
                    filtered_rows = filtered_rows[[0, 1]]
                    # Turning the value to negative
                    filtered_rows[1] = -abs(filtered_rows[1])

                    # Step 3: Combine the data
                    df_utanvetek_filtered.columns = range(len(df_utanvetek_filtered.columns))
                    df_utanvetek_filtered = pd.concat([df_utanvetek_filtered, filtered_rows], ignore_index=True)
                else:
                    # If "ÖSSZESÍTÉS" not found, add a warning
                    print(f"Warning: 'ÖSSZESÍTÉS' not found in file {file_path}")

                # Construct the output file path
                dir_name = os.path.dirname(file_path)
                base_name = os.path.basename(file_path).replace('.xlsx', '')
                output_file_path = os.path.join(dir_name, f"processed_{base_name}.xlsx")

                print(f"Saving to: {output_file_path}")

                # Save the DataFrame to a new Excel file with "processed_" prefix
                with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                    df_utanvetek_filtered.to_excel(writer, index=False, header=False, sheet_name='Sheet1')

                processed_files.append(output_file_path)
                print(f"Successfully saved: {output_file_path}")

            except Exception as e:
                error_msg = f"Error processing {file_path}: {str(e)}"
                errors.append(error_msg)
                print(error_msg)

        # Show appropriate message based on results
        if errors:
            error_text = "\n".join(errors)
            QMessageBox.critical(self.window, "Errors Occurred", f"The following errors occurred:\n{error_text}")
        elif processed_files:
            processed_text = "\n".join(processed_files)
            QMessageBox.information(self.window, "Success",
                                    f"All files processed successfully!\nFiles saved:\n{processed_text}")
        else:
            QMessageBox.warning(self.window, "No Files Processed", "No files were processed.")

    def go_back(self):
        # Show the main window again
        self.main_window.show()
        # Close this window
        self.window.close()

    def handle_close_event(self, event):
        # Show the main window again
        self.main_window.show()
        # Accept the event to close the window
        event.accept()