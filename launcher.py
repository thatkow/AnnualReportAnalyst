import tkinter as tk
from report_app import ReportAppV2


def main():
    """Start the ReportAppV2 GUI."""
    root = tk.Tk()
    app = ReportAppV2(root)
    root.mainloop()


if __name__ == "__main__":
    main()


