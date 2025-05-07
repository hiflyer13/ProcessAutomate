import tkinter as tk


class OTPWindow:
    def __init__(self, window, main_window):
        self.window = window
        self.main_window = main_window

        # Configure the window
        self.window.configure(bg="#f0f0f0")

        # Add title
        self.title_label = tk.Label(
            window,
            text="OTP",
            font=("Arial", 24, "bold"),
            bg="#f0f0f0"
        )
        self.title_label.pack(pady=30)

        # Content placeholder
        self.content_frame = tk.Frame(window, bg="#f0f0f0")
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=20)

        self.placeholder = tk.Label(
            self.content_frame,
            text="OTP processing functionality will be implemented here",
            font=("Arial", 14),
            bg="#f0f0f0"
        )
        self.placeholder.pack(pady=100)

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
        self.back_button.pack(pady=30)

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.go_back)

    def go_back(self):
        # Show the main window again
        self.main_window.deiconify()
        # Close this window
        self.window.destroy()