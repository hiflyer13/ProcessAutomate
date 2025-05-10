import os
import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QListWidget, QLineEdit, QFrame, \
    QMessageBox, QFileDialog, QHBoxLayout
import xlrd



class DPDWindow:
    
    def __init__(self, window, main_window):
        
        self.window = window
        self.main_window = main_window

        # Set up central widget and layout
        self.central_widget = QWidget()
        self.window.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Add title
        self.title_label = QLabel("DPD")
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

        # Optional label
        self.optional_label = QLabel("Optional")
        self.optional_label.setStyleSheet("font-size: 14px;")
        self.layout.addWidget(self.optional_label)

        # Frame for optional entry boxes
        self.optional_frame = QHBoxLayout()

        # First optional entry
        self.optional_entry1 = QLineEdit()
        self.optional_entry1.setPlaceholderText("Optional Entry 1")
        self.optional_frame.addWidget(self.optional_entry1)

        # Negative sign and second optional entry
        self.optional_entry2_frame = QHBoxLayout()
        self.negative_label = QLabel("-")
        self.negative_label.setStyleSheet("font-size: 12px;")
        self.optional_entry2_frame.addWidget(self.negative_label)

        self.optional_entry2 = QLineEdit()
        self.optional_entry2.setPlaceholderText("Positive integer only")
        self.optional_entry2.textChanged.connect(self.validate_positive_integer)
        self.optional_entry2_frame.addWidget(self.optional_entry2)

        self.optional_frame.addLayout(self.optional_entry2_frame)
        self.layout.addLayout(self.optional_frame)

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

    def validate_positive_integer(self, text):
        """Validate that the input is a positive integer or empty"""
        if text == "":
            return  # Allow empty field

        try:
            # Ensure it's a positive integer
            int_value = int(text)
            if int_value < 0:
                # Remove negative sign if user tries to enter one
                self.optional_entry2.setText(text.replace('-', ''))
        except ValueError:
            # If not an integer, remove the last character
            self.optional_entry2.setText(text[:-1])

    def browse_files(self):
        # Open file dialog to select multiple files
        file_paths, _ = QFileDialog.getOpenFileNames(self.window, "Select Files")
        for file_path in file_paths:
            self.file_listbox.addItem(file_path)

    def run_function(self):
            # Import the xlrd library here if it wasn't found in the global scope
        try:
            import xlrd
        except ImportError:
            QMessageBox.critical(self.window, "Missing Library", "The xlrd library is required but couldn't be imported.")
            return
        # Get all files from the listbox
        all_files = [self.file_listbox.item(i).text() for i in range(self.file_listbox.count())]

        if not all_files:
            QMessageBox.warning(self.window, "No Files", "Please browse and add files to process.")
            return

        optional_1 = self.optional_entry1.text()
        optional_2 = self.optional_entry2.text()

        # Convert optional_2 to negative integer if it's not empty
        if optional_2:
            try:
                optional_2 = -int(optional_2)  # Make it negative
            except ValueError:
                QMessageBox.critical(self.window, "Invalid Input", "Optional field 2 must be a number.")
                return

        processed_files = []
        errors = []

        for file_path in all_files:
            try:
                # Import a specific Excel sheet (DPD uses xls instead of xlsx, therefore, the xlrd library is needed)
                # Open the workbook
                # Load and process each file
                print(f"Processing file: {file_path}")
                workbook = xlrd.open_workbook(file_path)

                # Select the sheet by name or index
                sheet_name = 'Sheet1'  # Replace with the name of your sheet
                sheet = workbook.sheet_by_name(sheet_name)

                # Read the data into a pandas DataFrame, skipping the first 3 rows
                data = pd.DataFrame(sheet.get_rows())[3:]

                # Reset the index
                data.reset_index(drop=True, inplace=True)

                # Keep only the columns with index 2 and 5
                filtered_data = data[[2, 5]]

                # Reset the index
                filtered_data.reset_index(drop=True, inplace=True)

                # Ensure the column is treated as strings
                filtered_data[2] = filtered_data[2].astype(str)

                # Remove the "number:" string and strip spaces in column 2
                filtered_data.loc[:, 2] = filtered_data[2].str.replace('number:', '', regex=False).str.strip()  # Remove "number:" and strip spaces

                filtered_data.loc[:, 2] = filtered_data[2].astype(str).str.replace('.0', '', regex=False).astype(int)

                # Create a completely new column for the integer values
                filtered_data['integer_column'] = filtered_data[2].astype(str).str.replace('.0', '', regex=False)
                # Now convert to integers
                filtered_data['integer_column'] = pd.to_numeric(filtered_data['integer_column'])
                # Replace the original column
                filtered_data[2] = filtered_data['integer_column']

                filtered_data = filtered_data.drop('integer_column', axis=1)

                # First convert column 5 to string type
                filtered_data[5] = filtered_data[5].astype(str)

                # Then apply the transformations
                filtered_data[5] = filtered_data[5].str.replace("text:'", "", regex=False).str.split(" / ", expand=True)[0]

                filtered_data = filtered_data[[5, 2]]

                # Append optional entries if provided
                if optional_1 or optional_2:
                    optional_data = [optional_1, optional_2]
                    filtered_data.loc[filtered_data.shape[0]] = optional_data

                # Construct the output file path
                dir_name = os.path.dirname(file_path)
                base_name = os.path.basename(file_path).replace('.xls', '')
                output_file_path = os.path.join(dir_name, f"processed_{base_name}.xlsx")

                print(f"Saving to: {output_file_path}")

                # Save the DataFrame to a new Excel file with "processed_" prefix
                with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                    filtered_data.to_excel(writer, index=False, header=False, sheet_name='Sheet1')

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