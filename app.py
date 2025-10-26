"""Application entry point for the Annual Report Analyst UI."""

from __future__ import annotations

import logging
import tkinter as tk

from annual_report_analyst.report_app import ReportApp


logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def main() -> None:
    """Launch the Tkinter-based Annual Report Analyst application."""
    root = tk.Tk()
    ReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
