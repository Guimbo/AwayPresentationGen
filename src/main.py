import tkinter as tk
from interface import PresentationGeneratorApp

def main():
    root = tk.Tk()
    app = PresentationGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
