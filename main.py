import tkinter as tk
from main_menu import MainMenu

if __name__ == "__main__":
    root = tk.Tk()
    root.title("ProcessAutomate")
    root.geometry("800x650")

    app = MainMenu(root)

    root.mainloop()