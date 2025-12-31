import tkinter as tk
from tkinter import ttk
from scr.gui.main_window import MPesaApp

# from src.gui.main_window import MPesaApp

def main():
    root = tk.Tk()
    
    # Apply theme
    style = ttk.Style()
    style.theme_use('clam')  # Options: 'clam', 'alt', 'default', 'classic'
    
    # Custom button style
    style.configure('Accent.TButton', font=('Arial', 10, 'bold'))
    
    app = MPesaApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()