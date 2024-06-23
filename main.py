import sys
from GUI import FinancialDataApp
from TUI import main as tui_main

if __name__ == "__main__":
    if len(sys.argv) > 1:
        tui_main()
    else:
        import tkinter as tk
        root = tk.Tk()
        app = FinancialDataApp(root)
        root.mainloop()

