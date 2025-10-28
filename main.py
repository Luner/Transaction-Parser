#!/usr/bin/env python3
"""
Transaction Parser - Main Entry Point

Parse and categorize bank/credit card transactions from CSV exports.
Supports multiple bank formats including Apple Card, Capital One, Chase, and more.
"""

import sys
import tkinter as tk
from tkinter import ttk

# Add src directory to Python path
sys.path.insert(0, 'src')

from transaction_parser.gui import TransactionParserGUI


def main():
    """Main entry point for the Transaction Parser application"""
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
