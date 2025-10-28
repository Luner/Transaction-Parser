"""Application data directory management"""

import os
import platform
from pathlib import Path


class UnsupportedPlatformError(Exception):
    """Raised when the application is run on an unsupported platform"""
    pass


def get_app_data_dir() -> Path:
    """
    Get the application data directory for the current platform.

    Currently only supports macOS. On macOS, returns:
        ~/Library/Application Support/TransactionParser

    Returns:
        Path: The application data directory

    Raises:
        UnsupportedPlatformError: If the platform is not macOS
    """
    system = platform.system()

    if system == "Darwin":  # macOS
        app_data_dir = Path.home() / "Library" / "Application Support" / "TransactionParser"
    else:
        # TODO: Add support for Windows and Linux
        raise UnsupportedPlatformError(
            f"Platform '{system}' is not currently supported. "
            "TODO: Add support for Windows and Linux platforms."
        )

    # Create the directory if it doesn't exist
    app_data_dir.mkdir(parents=True, exist_ok=True)

    return app_data_dir


def get_category_mappings_path() -> Path:
    """
    Get the path to the category_mappings.json file.

    Returns:
        Path: Full path to category_mappings.json in the app data directory

    Raises:
        UnsupportedPlatformError: If the platform is not macOS
    """
    return get_app_data_dir() / "category_mappings.json"
