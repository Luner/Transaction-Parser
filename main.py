#!/usr/bin/env python3
"""
Transaction Parser - Main Entry Point

Parse and categorize bank/credit card transactions from CSV exports.
Supports multiple bank formats including Apple Card, Capital One, Chase, and more.
"""

import sys
import tkinter as tk
from tkinter import ttk, messagebox

# Add src directory to Python path
sys.path.insert(0, 'src')

from transaction_parser.gui import TransactionParserGUI
from transaction_parser.utils.app_data import UnsupportedPlatformError, get_app_data_dir


def main():
    """Main entry point for the Transaction Parser application"""
    # Check platform support before initializing GUI
    try:
        app_data_dir = get_app_data_dir()
        print(f"Using application data directory: {app_data_dir}")
    except UnsupportedPlatformError as e:
        # Show error dialog if possible, otherwise print to console
        root = tk.Tk()
        root.withdraw()  # Hide main window
        messagebox.showerror(
            "Unsupported Platform",
            str(e)
        )
        root.destroy()
        sys.exit(1)

    root = tk.Tk()

    # Set up styling
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('Accent.TButton', font=('Arial', 10, 'bold'))

    # Create and run the application
    app = TransactionParserGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
