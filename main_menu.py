import tkinter as tk
import importlib


class MainMenu:
    def __init__(self, master):
        self.master = master
        self.master.configure(bg="#f0f0f0")

        # Application title
        self.title_frame = tk.Frame(master, bg="#f0f0f0")
        self.title_frame.pack(pady=30)

        self.title_label = tk.Label(
            self.title_frame,
            text="ProcessAutomate",
            font=("Arial", 24, "bold"),
            bg="#f0f0f0"
        )
        self.title_label.pack()

        # Buttons frame
        self.button_frame = tk.Frame(master, bg="#f0f0f0")
        self.button_frame.pack(pady=20)

        # Button styling
        button_width = 20
        button_height = 2
        button_font = ("Arial", 12)
        button_bg = "#4a7abc"
        button_fg = "white"
        pady_between = 15

        # Create buttons for each module
        self.buttons = []
        modules = ["DPD", "Foxpost", "GLS", "MPL", "OTP", "Simple Pay"]

        for module in modules:
            button = tk.Button(
                self.button_frame,
                text=module,
                width=button_width,
                height=button_height,
                font=button_font,
                bg=button_bg,
                fg=button_fg,
                command=lambda m=module: self.open_module(m)
            )
            button.pack(pady=pady_between)
            self.buttons.append(button)

    def open_module(self, module_name):
        # Convert module name to lowercase for file naming
        module_file = module_name.lower().replace(" ", "_") + "_window"

        try:
            # Dynamically import the module
            module = importlib.import_module(module_file)

            # Hide main window
            self.master.withdraw()

            # Create a new window
            module_window = tk.Toplevel(self.master)
            module_window.title(f"ProcessAutomate - {module_name}")
            module_window.geometry("800x600")

            # Initialize the module's window class
            window_class = getattr(module, f"{module_name.replace(' ', '')}Window")
            window_class(module_window, self.master)

        except (ImportError, AttributeError) as e:
            print(f"Error loading module {module_name}: {e}")