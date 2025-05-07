import importlib
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QFrame, QSpacerItem, QSizePolicy

class MainMenu(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ProcessAutomate")
        self.setGeometry(100, 100, 800, 650)

        # Set up the main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setAlignment(Qt.AlignCenter)  # Center the layout

        # Application title
        self.title_frame = QFrame(self.central_widget)
        self.layout.addWidget(self.title_frame)

        self.title_label = QLabel("ProcessAutomate")
        self.title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.title_frame.setLayout(QVBoxLayout())
        self.title_frame.layout().addWidget(self.title_label)

        # Spacer to push buttons down
        self.layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Button styling
        button_width = 200
        button_height = 50
        button_font = "Arial"
        button_bg = "#4a7abc"
        button_fg = "white"
        pady_between = 15

        # Create buttons for each module
        self.buttons = []
        modules = ["DPD", "Foxpost", "GLS", "MPL", "OTP", "Simple Pay"]

        for module in modules:
            button = QPushButton(module)
            button.setFixedSize(button_width, button_height)
            button.setStyleSheet(f"background-color: {button_bg}; color: {button_fg}; font: {button_font};")
            button.clicked.connect(lambda checked, m=module: self.open_module(m))
            self.layout.addWidget(button)
            self.buttons.append(button)

            # Add a spacer between buttons
            self.layout.addSpacerItem(QSpacerItem(20, pady_between, QSizePolicy.Minimum, QSizePolicy.Fixed))

        # Final spacer to push buttons up
        self.layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

    def open_module(self, module_name):
        # Convert module name to lowercase for file naming
        module_file = module_name.lower().replace(" ", "_") + "_window"

        try:
            # Dynamically import the module
            module = importlib.import_module(module_file)

            # Hide main window
            self.hide()

            # Create a new window
            module_window = QMainWindow()
            module_window.setWindowTitle(f"ProcessAutomate - {module_name}")
            module_window.setGeometry(100, 100, 800, 600)

            # Initialize the module's window class
            window_class = getattr(module, f"{module_name.replace(' ', '')}Window")
            window_class(module_window, self)

            module_window.show()

        except (ImportError, AttributeError) as e:
            print(f"Error loading module {module_name}: {e}")