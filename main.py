import tkinter as tk
from ui import MainWindow

def main():
    root = tk.Tk()
    app = MainWindow(root)
    
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.mainloop()

if __name__ == "__main__":
    main()
