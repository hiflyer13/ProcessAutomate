import os
import pandas as pd
import xml.etree.ElementTree as ET
import re
from PyQt5.QtWidgets import (QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
                            QWidget, QMessageBox, QLineEdit, QFileDialog, 
                            QListWidget, QSizePolicy, QTextEdit)
from PyQt5.QtCore import QThread, pyqtSignal, Qt

class ProcessingThread(QThread):
    """Thread for processing files"""
    progress_update = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, xml_path, file_type, files):
        super().__init__()
        self.xml_path = xml_path
        self.file_type = file_type
        self.files = files
        
    def run(self):
        try:
            # Process the XML file first - this is required for all file types
            self.progress_update.emit(f"Starting {self.file_type} file processing...")
            self.progress_update.emit(f"Loading XML file: {self.xml_path}")
            
            # Process XML file to get reference data
            df_xml = self.process_xml_file(self.xml_path)
            
            if df_xml is None or df_xml.empty:
                self.progress_update.emit("Error: Failed to process XML file or no valid data found")
                self.finished.emit(False, "XML processing failed")
                return
                
            self.progress_update.emit(f"XML processing complete. Found {len(df_xml)} reference records.")
            
            # Process the selected files based on type
            if not self.files:
                self.progress_update.emit(f"No {self.file_type} files selected.")
                self.finished.emit(False, "No files to process")
                return
                
            # Process files based on type
            processed_count = 0
            for file_path in self.files:
                try:
                    self.progress_update.emit(f"Processing file: {os.path.basename(file_path)}")
                    
                    if self.file_type == "equal":
                        self.process_equal_file(file_path, df_xml)
                    elif self.file_type == "pg":
                        self.process_pg_file(file_path, df_xml)
                    elif self.file_type == "t":
                        self.process_t_file(file_path, df_xml)
                        
                    processed_count += 1
                    self.progress_update.emit(f"✓ Successfully processed: {os.path.basename(file_path)}")
                except Exception as e:
                    self.progress_update.emit(f"✗ Error processing {os.path.basename(file_path)}: {str(e)}")
            
            if processed_count > 0:
                self.finished.emit(True, f"Successfully processed {processed_count} {self.file_type} files")
            else:
                self.finished.emit(False, f"No {self.file_type} files were processed successfully")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.progress_update.emit(f"Error: {str(e)}")
            self.progress_update.emit(error_details)
            self.finished.emit(False, f"Error during {self.file_type} processing: {str(e)}")
    
    def process_xml_file(self, xml_path):
        """Process the XML file and extract reference data"""
        try:
            # Define namespaces
            namespaces = {
                'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
                'o': 'urn:schemas-microsoft-com:office:office',
                'x': 'urn:schemas-microsoft-com:office:excel',
            }

            self.progress_update.emit("Parsing XML file...")
            tree = ET.parse(xml_path)
            root = tree.getroot()

            # Find all Worksheet elements
            worksheets = root.findall('.//ss:Worksheet', namespaces)
            
            if not worksheets:
                self.progress_update.emit("Error: No worksheets found in XML")
                return pd.DataFrame({'Sorszám': [], 'Hivatkozás': []})
            
            self.progress_update.emit(f"Found {len(worksheets)} worksheets in XML")
            
            # Process each worksheet to find one with the right data
            for worksheet in worksheets:
                worksheet_name = worksheet.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', "Unnamed")
                self.progress_update.emit(f"Checking worksheet: {worksheet_name}")
                
                # Find the Table element within the Worksheet
                table = worksheet.find('.//ss:Table', namespaces)
                if table is None:
                    continue
                    
                # Extract rows
                rows = table.findall('./ss:Row', namespaces)
                if not rows:
                    continue
                    
                # Process rows to extract data
                data = []
                for row in rows:
                    row_data = []
                    cells = row.findall('./ss:Cell', namespaces)
                    
                    for cell in cells:
                        # Handle merged cells
                        index_attr = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                        if index_attr:
                            current_index = len(row_data) + 1
                            for _ in range(int(index_attr) - current_index):
                                row_data.append(None)
                        
                        # Get cell value
                        data_element = cell.find('./ss:Data', namespaces)
                        cell_value = data_element.text if data_element is not None else None
                        row_data.append(cell_value)
                    
                    data.append(row_data)
                
                # Check if we have enough data
                if len(data) <= 1:
                    continue
                    
                # Get headers from the first row
                headers = data[0]
                headers = [str(h) if h is not None else f"Column_{i}" for i, h in enumerate(headers)]
                
                # Create DataFrame
                df = pd.DataFrame(data[1:], columns=headers)
                
                # Check if this is the worksheet we want (has required columns)
                if 'Sorszám' in df.columns and 'Hivatkozás' in df.columns:
                    self.progress_update.emit(f"Found required columns in worksheet {worksheet_name}")
                    
                    # Extract reference data
                    df_filtered = df[['Sorszám', 'Hivatkozás']]
                    
                    # Extract trailing numbers from Hivatkozás
                    df_filtered['Hivatkozás'] = df_filtered['Hivatkozás'].apply(
                        lambda x: self.extract_trailing_numbers(x) if not pd.isna(x) else ""
                    )
                    
                    # Clean data
                    df_filtered = df_filtered.dropna()
                    df_filtered = df_filtered[df_filtered['Hivatkozás'] != ""]
                    
                    return df_filtered
            
            self.progress_update.emit("Could not find a worksheet with the required columns")
            return pd.DataFrame({'Sorszám': [], 'Hivatkozás': []})
            
        except Exception as e:
            self.progress_update.emit(f"Error processing XML file: {str(e)}")
            import traceback
            self.progress_update.emit(traceback.format_exc())
            return pd.DataFrame({'Sorszám': [], 'Hivatkozás': []})
    
    def extract_trailing_numbers(self, text):
        """Extract trailing numeric characters from text"""
        if pd.isna(text):
            return ""
        
        text = str(text)
        match = re.search(r'(\d+)$', text)
        if match:
            return match.group(1)
        else:
            return ""
    
    def process_equal_file(self, file_path, df_xml):
        """Process an Equal Sign (=) file"""
        self.progress_update.emit(f"Reading Equal file: {os.path.basename(file_path)}")
        
        # Read input file
        if file_path.endswith('.csv'):
            df_equal = pd.read_csv(file_path, sep=';')
        else:  # Excel file
            df_equal = pd.read_excel(file_path)
        
        # Verify required columns
        required_columns = ["Tranzakciós jutalék", "Kereskedői tranzakció ID", "Tranzakció összege"]
        missing_columns = [col for col in required_columns if col not in df_equal.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Process transaction fees
        self.progress_update.emit("Processing transaction fees...")
        df_equal["Tranzakciós jutalék"] = df_equal["Tranzakciós jutalék"].str.replace(",00", "").astype(int)
        sum_equal_jutalek = df_equal["Tranzakciós jutalék"].sum()
        
        # Process transaction amounts
        self.progress_update.emit("Processing transaction amounts...")
        filtered_df = df_equal[["Kereskedői tranzakció ID", "Tranzakció összege"]]

        filtered_df["Tranzakció összege"] = filtered_df["Tranzakció összege"].str.replace(",00", "").astype(int)

        
        # Clean transaction IDs
        filtered_df["Kereskedői tranzakció ID"] = (filtered_df["Kereskedői tranzakció ID"]
                                                .astype(str)
                                                .str.replace('="', '', regex=False)
                                                .str.replace('"', '', regex=False))

        
        # Create mapping from XML data
        self.progress_update.emit("Mapping transaction IDs to reference numbers...")
        mapping_dict = dict(zip(df_xml['Hivatkozás'], df_xml['Sorszám']))
        
        # Apply mapping to get Sorszám
        final_df = filtered_df.copy()
        final_df['Sorszám'] = final_df['Kereskedői tranzakció ID'].map(mapping_dict)
        final_df['Sorszám'] = final_df['Sorszám'].fillna("1")  # Default value if not found
        
        # Reorder columns
        final_df = final_df[['Sorszám', 'Tranzakció összege']]
        
        # Add fee row
        new_row = pd.DataFrame({
            'Sorszám': ['1'],
            'Tranzakció összege': [-abs(sum_equal_jutalek)]  # Make negative
        })
        
        final_df = pd.concat([final_df, new_row], ignore_index=True)
        
        # Save result
        output_path = os.path.join(
            os.path.dirname(file_path),
            f"processed_{os.path.basename(file_path)}"
        )

        # Ensure output has .xlsx extension
        if not output_path.endswith('.xlsx'):
            output_path = os.path.splitext(output_path)[0] + '.xlsx'

        self.progress_update.emit(f"Saving result to {os.path.basename(output_path)}...")
        final_df.to_excel(output_path, index=False, header=False)


        # We are going to duplicate the code here as we need an extended Excel file.
        filtered_df_extended = df_equal[["Kereskedői tranzakció ID", "Tranzakció összege", "Vásárló", "E-mail cím"]]
        filtered_df_extended["Tranzakció összege"] = filtered_df_extended["Tranzakció összege"].str.replace(",00", "").astype(int)

        filtered_df_extended["Kereskedői tranzakció ID"] = (filtered_df_extended["Kereskedői tranzakció ID"]
                                                .astype(str)
                                                .str.replace('="', '', regex=False)
                                                .str.replace('"', '', regex=False))
        
        final_df_extended = filtered_df_extended.copy()
        final_df_extended['Sorszám'] = final_df_extended['Kereskedői tranzakció ID'].map(mapping_dict)
        final_df_extended['Sorszám'] = final_df_extended['Sorszám'].fillna("1")  # Default value if not found
        final_df_extended = final_df_extended[['Sorszám', 'Tranzakció összege', 'Vásárló', 'E-mail cím']]
        final_df_extended = pd.concat([final_df_extended, new_row], ignore_index=True)


        output_path_extended = os.path.join(
            os.path.dirname(file_path),
            f"processed_extended_{os.path.basename(file_path)}"
        )
        
        if not output_path_extended.endswith('.xlsx'):
            output_path_extended = os.path.splitext(output_path_extended)[0] + '.xlsx'
        

        final_df_extended.to_excel(output_path_extended, index=False, header=False)
        
        return output_path
    
    def process_pg_file(self, file_path, df_xml):
        """Process a PG file"""
        self.progress_update.emit(f"Reading PG file: {os.path.basename(file_path)}")
        
        # Read input file
        if file_path.endswith('.csv'):
            df_pg = pd.read_csv(file_path, sep=';')
        else:  # Excel file
            df_pg = pd.read_excel(file_path)
        
        # Verify required columns
        required_columns = ["Tranzakciós jutalék", "Kereskedői tranzakció ID", "Tranzakció összege"]
        missing_columns = [col for col in required_columns if col not in df_pg.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Process transaction fees
        self.progress_update.emit("Processing transaction fees...")
        df_pg["Tranzakciós jutalék"] = df_pg["Tranzakciós jutalék"].str.replace(",00", "").astype(int)
        sum_pg_jutalek = df_pg["Tranzakciós jutalék"].sum()
        
        # Process transaction amounts
        self.progress_update.emit("Processing transaction amounts...")
        filtered_df = df_pg[["Kereskedői tranzakció ID", "Tranzakció összege"]]

        filtered_df["Tranzakció összege"] = filtered_df["Tranzakció összege"].str.replace(",00", "").astype(int)
        
        # Clean transaction IDs - for PG files, we need to remove the "pg-" prefix
        filtered_df["Kereskedői tranzakció ID"] = filtered_df["Kereskedői tranzakció ID"].astype(str).str.replace('pg-', '', regex=False)
        
        # Create mapping from XML data
        self.progress_update.emit("Mapping transaction IDs to reference numbers...")
        mapping_dict = dict(zip(df_xml['Hivatkozás'], df_xml['Sorszám']))
        
        # Apply mapping to get Sorszám
        final_df = filtered_df.copy()
        final_df['Sorszám'] = final_df['Kereskedői tranzakció ID'].map(mapping_dict)
        final_df['Sorszám'] = final_df['Sorszám'].fillna("1")  # Default value if not found
        
        # Reorder columns
        final_df = final_df[['Sorszám', 'Tranzakció összege']]
        
        # Add fee row
        new_row = pd.DataFrame({
            'Sorszám': ['1'],
            'Tranzakció összege': [-abs(sum_pg_jutalek)]  # Make negative
        })
        
        final_df = pd.concat([final_df, new_row], ignore_index=True)
        
        # Save result
        output_path = os.path.join(
            os.path.dirname(file_path),
            f"processed_{os.path.basename(file_path)}"
        )
        
        # Ensure output has .xlsx extension
        if not output_path.endswith('.xlsx'):
            output_path = os.path.splitext(output_path)[0] + '.xlsx'
        
        self.progress_update.emit(f"Saving result to {os.path.basename(output_path)}...")
        final_df.to_excel(output_path, index=False, header=False)

        # We are going to duplicate the code here as we need an extended Excel file.
        filtered_df_extended = df_pg[["Kereskedői tranzakció ID", "Tranzakció összege", "Vásárló", "E-mail cím"]]
        filtered_df_extended["Tranzakció összege"] = filtered_df_extended["Tranzakció összege"].str.replace(",00", "").astype(int)
        filtered_df_extended["Kereskedői tranzakció ID"] = filtered_df_extended["Kereskedői tranzakció ID"].astype(str).str.replace('pg-', '', regex=False)
        final_df_extended = filtered_df_extended.copy()
        final_df_extended['Sorszám'] = final_df_extended['Kereskedői tranzakció ID'].map(mapping_dict)
        final_df_extended['Sorszám'] = final_df_extended['Sorszám'].fillna("1")  # Default value if not found
        final_df_extended = final_df_extended[['Sorszám', 'Tranzakció összege', 'Vásárló', 'E-mail cím']]
        final_df_extended = pd.concat([final_df_extended, new_row], ignore_index=True)
        output_path_extended = os.path.join(
            os.path.dirname(file_path),
            f"processed_extended_{os.path.basename(file_path)}"
        )
        if not output_path_extended.endswith('.xlsx'):
            output_path_extended = os.path.splitext(output_path_extended)[0] + '.xlsx'
        
        final_df_extended.to_excel(output_path_extended, index=False, header=False)
        
        return output_path
    
    def process_t_file(self, file_path, df_xml):
        """Process a T file"""
        self.progress_update.emit(f"Reading T file: {os.path.basename(file_path)}")
        
        # Read input file
        if file_path.endswith('.csv'):
            df_t = pd.read_csv(file_path, sep=';')
        else:  # Excel file
            df_t = pd.read_excel(file_path)
        
        # Verify required columns
        required_columns = ["Tranzakciós jutalék", "Kereskedői tranzakció ID", "Tranzakció összege"]
        missing_columns = [col for col in required_columns if col not in df_t.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Process transaction fees
        self.progress_update.emit("Processing transaction fees...")
        df_t["Tranzakciós jutalék"] = df_t["Tranzakciós jutalék"].str.replace(",00", "").astype(int)
        sum_t_jutalek = df_t["Tranzakciós jutalék"].sum()
        
        # Process transaction amounts
        self.progress_update.emit("Processing transaction amounts...")
        filtered_df = df_t[["Kereskedői tranzakció ID", "Tranzakció összege"]]
        filtered_df["Tranzakció összege"] = filtered_df["Tranzakció összege"].str.replace(",00", "").astype(int)
        
        # Clean transaction IDs - for T files, we split at "T" and take the second part
        self.progress_update.emit("Processing T-type transaction IDs...")
        filtered_df["Kereskedői tranzakció ID"] = filtered_df["Kereskedői tranzakció ID"].astype(str)
        
        # Safe splitting at T
        def split_at_t(id_string):
            parts = id_string.split("T")
            if len(parts) > 1:
                return parts[1]
            else:
                return id_string
                
        filtered_df["Kereskedői tranzakció ID"] = filtered_df["Kereskedői tranzakció ID"].apply(split_at_t)
        
        # Create mapping from XML data
        self.progress_update.emit("Mapping transaction IDs to reference numbers...")
        mapping_dict = dict(zip(df_xml['Hivatkozás'], df_xml['Sorszám']))
        
        # Apply mapping to get Sorszám
        final_df = filtered_df.copy()
        final_df['Sorszám'] = final_df['Kereskedői tranzakció ID'].map(mapping_dict)
        final_df['Sorszám'] = final_df['Sorszám'].fillna("1")  # Default value if not found
        
        # Reorder columns
        final_df = final_df[['Sorszám', 'Tranzakció összege']]
        
        # Add fee row
        new_row = pd.DataFrame({
            'Sorszám': ['1'],
            'Tranzakció összege': [-abs(sum_t_jutalek)]  # Make negative
        })
        
        final_df = pd.concat([final_df, new_row], ignore_index=True)
        
        # Save result
        output_path = os.path.join(
            os.path.dirname(file_path),
            f"processed_{os.path.basename(file_path)}"
        )
        
        # Ensure output has .xlsx extension
        if not output_path.endswith('.xlsx'):
            output_path = os.path.splitext(output_path)[0] + '.xlsx'
        
        self.progress_update.emit(f"Saving result to {os.path.basename(output_path)}...")
        final_df.to_excel(output_path, index=False, header=False)

        # We are going to duplicate the code here as we need an extended Excel file.
        filtered_df_extended = df_t[["Kereskedői tranzakció ID", "Tranzakció összege", "Vásárló", "E-mail cím"]]
        filtered_df_extended["Tranzakció összege"] = filtered_df_extended["Tranzakció összege"].str.replace(",00", "").astype(int)
        filtered_df_extended["Kereskedői tranzakció ID"] = filtered_df_extended["Kereskedői tranzakció ID"].astype(str)
        filtered_df_extended["Kereskedői tranzakció ID"] = filtered_df_extended["Kereskedői tranzakció ID"].apply(split_at_t)
        final_df_extended = filtered_df_extended.copy()
        final_df_extended['Sorszám'] = final_df_extended['Kereskedői tranzakció ID'].map(mapping_dict)
        final_df_extended['Sorszám'] = final_df_extended['Sorszám'].fillna("1")  # Default value if not found
        final_df_extended = final_df_extended[['Sorszám', 'Tranzakció összege', 'Vásárló', 'E-mail cím']]
        final_df_extended = pd.concat([final_df_extended, new_row], ignore_index=True)
        output_path_extended = os.path.join(
            os.path.dirname(file_path),
            f"processed_extended_{os.path.basename(file_path)}"
        )
        if not output_path_extended.endswith('.xlsx'):
            output_path_extended = os.path.splitext(output_path_extended)[0] + '.xlsx'
        final_df_extended.to_excel(output_path_extended, index=False, header=False)
        
        return output_path


class SimplePayWindow:
    
    def __init__(self, window, main_window):
        
        self.window = window
        self.main_window = main_window
        self.processing_thread = None

        # Set up central widget and layout
        self.central_widget = QWidget()
        self.window.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Add title
        self.title_label = QLabel("Simple Pay")
        self.title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.layout.addWidget(self.title_label)
        
        # Add XML file selection row
        self.file_layout = QHBoxLayout()
        
        # XML label
        self.xml_label = QLabel("XML:")
        self.xml_label.setStyleSheet("font-size: 14px;")
        self.file_layout.addWidget(self.xml_label)
        
        # File path entry box
        self.file_path_entry = QLineEdit()
        self.file_path_entry.setMinimumWidth(300)
        self.file_layout.addWidget(self.file_path_entry)
        
        # Browse button
        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_file)
        self.file_layout.addWidget(self.browse_button)
        
        # Add file selection row to main layout
        self.layout.addLayout(self.file_layout)
        
        # Create layout for the 3 list boxes
        self.lists_layout = QHBoxLayout()
        
        # Create the three columns for lists
        # Equal Sign column
        self.equal_layout = QVBoxLayout()
        self.equal_label = QLabel("Equal Sign =")
        self.equal_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.equal_layout.addWidget(self.equal_label)
        
        self.equal_list = QListWidget()
        self.equal_list.setMinimumHeight(300)
        self.equal_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.equal_layout.addWidget(self.equal_list)
        
        # Equal Sign buttons layout
        self.equal_buttons_layout = QHBoxLayout()
        
        # Equal Sign Browse button
        self.equal_browse = QPushButton("Browse")
        self.equal_browse.clicked.connect(self.browse_equal_files)
        self.equal_buttons_layout.addWidget(self.equal_browse)
        
        # Equal Sign Clear button
        self.equal_clear = QPushButton("Clear")
        self.equal_clear.clicked.connect(self.equal_list.clear)
        self.equal_buttons_layout.addWidget(self.equal_clear)
        
        # Equal Sign Run button
        self.equal_run = QPushButton("Run Equal")
        self.equal_run.setStyleSheet("background-color: #4a7abc; color: white;")
        self.equal_run.clicked.connect(lambda: self.run_files("equal"))
        self.equal_buttons_layout.addWidget(self.equal_run)
        
        # Add buttons to layout
        self.equal_layout.addLayout(self.equal_buttons_layout)
        
        self.lists_layout.addLayout(self.equal_layout)
        
        # PG column
        self.pg_layout = QVBoxLayout()
        self.pg_label = QLabel("PG")
        self.pg_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.pg_layout.addWidget(self.pg_label)
        
        self.pg_list = QListWidget()
        self.pg_list.setMinimumHeight(300)
        self.pg_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.pg_layout.addWidget(self.pg_list)
        
        # PG buttons layout
        self.pg_buttons_layout = QHBoxLayout()
        
        # PG Browse button
        self.pg_browse = QPushButton("Browse")
        self.pg_browse.clicked.connect(self.browse_pg_files)
        self.pg_buttons_layout.addWidget(self.pg_browse)
        
        # PG Clear button
        self.pg_clear = QPushButton("Clear") 
        self.pg_clear.clicked.connect(self.pg_list.clear)
        self.pg_buttons_layout.addWidget(self.pg_clear)
        
        # PG Run button
        self.pg_run = QPushButton("Run PG")
        self.pg_run.setStyleSheet("background-color: #4a7abc; color: white;")
        self.pg_run.clicked.connect(lambda: self.run_files("pg"))
        self.pg_buttons_layout.addWidget(self.pg_run)
        
        # Add buttons to layout
        self.pg_layout.addLayout(self.pg_buttons_layout)
        
        self.lists_layout.addLayout(self.pg_layout)
        
        # T column
        self.t_layout = QVBoxLayout()
        self.t_label = QLabel("T")
        self.t_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.t_layout.addWidget(self.t_label)
        
        self.t_list = QListWidget()
        self.t_list.setMinimumHeight(300)
        self.t_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.t_layout.addWidget(self.t_list)
        
        # T buttons layout
        self.t_buttons_layout = QHBoxLayout()
        
        # T Browse button
        self.t_browse = QPushButton("Browse")
        self.t_browse.clicked.connect(self.browse_t_files)
        self.t_buttons_layout.addWidget(self.t_browse)
        
        # T Clear button
        self.t_clear = QPushButton("Clear")
        self.t_clear.clicked.connect(self.t_list.clear)
        self.t_buttons_layout.addWidget(self.t_clear)
        
        # T Run button
        self.t_run = QPushButton("Run T")
        self.t_run.setStyleSheet("background-color: #4a7abc; color: white;")
        self.t_run.clicked.connect(lambda: self.run_files("t"))
        self.t_buttons_layout.addWidget(self.t_run)
        
        # Add buttons to layout
        self.t_layout.addLayout(self.t_buttons_layout)
        
        self.lists_layout.addLayout(self.t_layout)
        
        # Add the lists layout to the main layout
        self.layout.addLayout(self.lists_layout, 1)
        
        # Add log display area
        self.log_label = QLabel("Processing Log:")
        self.log_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.layout.addWidget(self.log_label)
        
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMinimumHeight(150)
        self.layout.addWidget(self.log_display)
        
        # Back button
        self.back_button = QPushButton("Back to Main Menu")
        self.back_button.setStyleSheet("background-color: #4a7abc; color: white; font-size: 12px; padding: 10px;")
        self.back_button.clicked.connect(self.go_back)
        self.layout.addWidget(self.back_button)

        # Connect close event
        self.window.closeEvent = self.handle_close_event
    
    def browse_file(self):
        """Open a file dialog to browse for an XML file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self.window,
            "Select XML File",
            "",
            "XML Files (*.xml);;All Files (*)"
        )
        if file_path:
            self.file_path_entry.setText(file_path)
            self.update_progress(f"Selected XML file: {file_path}")
    
    def browse_equal_files(self):
        """Open a file dialog to browse for Equal Sign files"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self.window,
            "Select Equal Sign Files",
            "",
            "CSV/Excel Files (*.csv *.xlsx *.xls);;All Files (*)"
        )
        
        if file_paths:
            self.update_progress(f"Selected {len(file_paths)} Equal Sign files")
            for file_path in file_paths:
                self.equal_list.addItem(file_path)
                
    def browse_pg_files(self):
        """Open a file dialog to browse for PG files"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self.window,
            "Select PG Files",
            "",
            "CSV/Excel Files (*.csv *.xlsx *.xls);;All Files (*)"
        )
        
        if file_paths:
            self.update_progress(f"Selected {len(file_paths)} PG files")
            for file_path in file_paths:
                self.pg_list.addItem(file_path)
    
    def browse_t_files(self):
        """Open a file dialog to browse for T files"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self.window,
            "Select T Files",
            "",
            "CSV/Excel Files (*.csv *.xlsx *.xls);;All Files (*)"
        )
        
        if file_paths:
            self.update_progress(f"Selected {len(file_paths)} T files")
            for file_path in file_paths:
                self.t_list.addItem(file_path)

    def run_files(self, file_type):
        """Process files of a specific type"""
        xml_path = self.file_path_entry.text()
        
        # Validate XML file
        if not xml_path or not os.path.isfile(xml_path):
            QMessageBox.warning(self.window, "Warning", "Please select a valid XML file")
            return
        
        # Get files based on type
        files = []
        if file_type == "equal":
            files = [self.equal_list.item(i).text() for i in range(self.equal_list.count())]
            self.equal_run.setEnabled(False)
            self.equal_run.setText("Processing...")
        elif file_type == "pg":
            files = [self.pg_list.item(i).text() for i in range(self.pg_list.count())]
            self.pg_run.setEnabled(False)
            self.pg_run.setText("Processing...")
        elif file_type == "t":
            files = [self.t_list.item(i).text() for i in range(self.t_list.count())]
            self.t_run.setEnabled(False)
            self.t_run.setText("Processing...")
        
        # Validate that files were selected
        if not files:
            QMessageBox.warning(self.window, "Warning", f"Please select at least one {file_type} file to process")
            self.reset_run_button(file_type)
            return
        
        # Create and start the processing thread
        self.log_display.clear()  # Clear log before starting new process
        self.processing_thread = ProcessingThread(xml_path, file_type, files)
        self.processing_thread.progress_update.connect(self.update_progress)
        self.processing_thread.finished.connect(lambda success, msg: self.processing_finished(success, msg, file_type))
        self.processing_thread.start()
    
    def reset_run_button(self, file_type):
        """Reset the run button for a specific file type"""
        if file_type == "equal":
            self.equal_run.setEnabled(True)
            self.equal_run.setText("Run Equal")
        elif file_type == "pg":
            self.pg_run.setEnabled(True)
            self.pg_run.setText("Run PG")
        elif file_type == "t":
            self.t_run.setEnabled(True)
            self.t_run.setText("Run T")
    
    def update_progress(self, message):
        """Update the user interface with progress messages"""
        print(message)  # Print to console
        self.log_display.append(message)  # Also add to log display
        # Make sure the new text is visible
        scrollbar = self.log_display.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
        # Force the application to process events to keep UI responsive
        from PyQt5.QtCore import QCoreApplication
        QCoreApplication.processEvents()
    
    def processing_finished(self, success, message, file_type):
        """Handle the completion of processing"""
        self.reset_run_button(file_type)
        
        if success:
            self.update_progress("✅ " + message)
            QMessageBox.information(self.window, "Success", message)
        else:
            self.update_progress("❌ " + message)
            QMessageBox.critical(self.window, "Error", message)
        
    def go_back(self):
        """Return to the main menu"""
        if self.processing_thread and self.processing_thread.isRunning():
            QMessageBox.warning(self.window, "Warning", "Please wait for processing to complete.")
            return
            
        self.main_window.show()
        self.window.close()

    def handle_close_event(self, event):
        """Handle window close event"""
        if self.processing_thread and self.processing_thread.isRunning():
            QMessageBox.warning(self.window, "Warning", "Please wait for processing to complete.")
            event.ignore()
            return
            
        self.main_window.show()
        event.accept()