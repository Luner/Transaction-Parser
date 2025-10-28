"""
Bank and credit card CSV format configurations.

This module contains predefined configurations for various banks and credit cards,
making it easy to import transactions without manually entering column names.
"""

from typing import Dict, Optional


class BankFormat:
    """Represents a bank/card CSV format configuration"""

    def __init__(self, name: str, date_col: str, desc_col: str,
                 amount_col: str = None, debit_col: str = None,
                 credit_col: str = None, date_format: str = "%m/%d/%Y",
                 invert_amounts: bool = False, description: str = "",
                 has_header: bool = True):
        self.name = name
        self.date_col = date_col
        self.desc_col = desc_col
        self.amount_col = amount_col
        self.debit_col = debit_col
        self.credit_col = credit_col
        self.date_format = date_format
        self.invert_amounts = invert_amounts
        self.description = description
        self.has_header = has_header  # False for headerless CSVs like Wells Fargo Checking

    def to_dict(self) -> Dict:
        """Convert to dictionary representation"""
        return {
            'name': self.name,
            'date_col': self.date_col,
            'desc_col': self.desc_col,
            'amount_col': self.amount_col,
            'debit_col': self.debit_col,
            'credit_col': self.credit_col,
            'date_format': self.date_format,
            'invert_amounts': self.invert_amounts,
            'description': self.description
        }


# Predefined bank/card formats
BANK_FORMATS = {
    'apple_card': BankFormat(
        name='Apple Card',
        date_col='Transaction Date',
        desc_col='Merchant',
        amount_col='Amount (USD)',
        date_format='%m/%d/%Y',
        invert_amounts=True,
        description='Apple Card CSV export format'
    ),

    'capital_one': BankFormat(
        name='Capital One',
        date_col='Transaction Date',
        desc_col='Description',
        debit_col='Debit',
        credit_col='Credit',
        date_format='%m/%d/%Y',
        invert_amounts=False,
        description='Capital One CSV export with separate Debit/Credit columns'
    ),

    'chase': BankFormat(
        name='Chase',
        date_col='Transaction Date',
        desc_col='Description',
        amount_col='Amount',
        date_format='%m/%d/%Y',
        invert_amounts=False,
        description='Chase Bank CSV export format'
    ),

    'wells_fargo_bank': BankFormat(
        name='Wells Fargo Bank',
        date_col='0',  # Column index 0 for headerless CSV
        desc_col='4',  # Column index 4 for headerless CSV
        amount_col='1',  # Column index 1 for headerless CSV
        date_format='%m/%d/%Y',
        invert_amounts=False,
        has_header=False,
        description='Wells Fargo Checking - headerless CSV format'
    ),

    'custom': BankFormat(
        name='Custom',
        date_col='Date',
        desc_col='Description',
        amount_col='Amount',
        date_format='%m/%d/%Y',
        invert_amounts=False,
        description='Custom format - manually configure column names'
    )
}


def get_bank_format(format_key: str) -> Optional[BankFormat]:
    """Get a bank format by its key"""
    return BANK_FORMATS.get(format_key.lower())


def get_all_bank_names() -> list:
    """Get list of all available bank/card names"""
    return [fmt.name for fmt in BANK_FORMATS.values()]


def get_bank_format_by_name(name: str) -> Optional[BankFormat]:
    """Get a bank format by its display name"""
    for fmt in BANK_FORMATS.values():
        if fmt.name.lower() == name.lower():
            return fmt
    return None


def add_custom_format(key: str, format_config: BankFormat):
    """Add a custom bank format at runtime"""
    BANK_FORMATS[key] = format_config


def detect_bank_format_from_headers(headers: list) -> Optional[BankFormat]:
    """
    Detect bank format by matching CSV headers against known formats.
    Also handles headerless CSVs by validating data patterns.

    Args:
        headers: List of column headers (or first data row for headerless CSV)

    Returns:
        BankFormat object if a match is found, None otherwise
    """
    if not headers:
        return None

    # Normalize headers for comparison (strip whitespace)
    normalized_headers = [h.strip() for h in headers]

    # First, try header-based matching
    for fmt in BANK_FORMATS.values():
        # Skip Custom format and headerless formats in this pass
        if fmt.name == "Custom" or not fmt.has_header:
            continue

        # Check if all required columns exist in the CSV headers
        required_cols = [fmt.date_col, fmt.desc_col]

        # Add amount-related columns based on format type
        if fmt.amount_col:
            required_cols.append(fmt.amount_col)
        elif fmt.debit_col and fmt.credit_col:
            required_cols.extend([fmt.debit_col, fmt.credit_col])

        # Check if all required columns are present in headers
        if all(col in normalized_headers for col in required_cols):
            return fmt

    # If no header match, try headerless format detection
    # Check if this looks like Wells Fargo (date, amount, *, *, description pattern)
    if len(normalized_headers) >= 5:
        try:
            from datetime import datetime
            # CSV reader already removes quotes, so we don't need strip('"')
            # Try to parse first column as date
            date_str = normalized_headers[0]
            datetime.strptime(date_str, '%m/%d/%Y')

            # Try to parse second column as amount (negative or positive number)
            amount_str = normalized_headers[1].replace(',', '')
            float(amount_str)

            # If we got here, it looks like Wells Fargo Checking format
            wells_fargo = BANK_FORMATS.get('wells_fargo_bank')
            if wells_fargo:
                return wells_fargo
        except (ValueError, IndexError):
            # Silent fail - just return None if detection doesn't work
            pass

    return None
