import time

import xlwings as xw


def place_complete_limit_order():
    """Place a limit order with all required fields including Symbol and Exchange"""

    print("=" * 60)
    print("PLACING COMPLETE LIMIT ORDER")
    print("=" * 60)

    # Connect to Excel
    wb = xw.books.active

    # Find sheets
    data_sheet = None
    place_order_sheet = None

    for sheet in wb.sheets:
        if sheet.name == 'Charts':
            data_sheet = sheet
        elif 'Place Orders' in sheet.name:
            place_order_sheet = sheet

    if not place_order_sheet:
        print("✗ Place Orders sheet not found!")
        return False

    # Get current price from data sheet
    current_bid = 23547.00  # Default if can't read
    if data_sheet:
        bid = data_sheet.range('B2').value
        if bid:
            current_bid = bid
            print(f"Current Bid: {current_bid}")

    # Calculate safe limit price
    limit_price = round(current_bid - 100, 2)

    print(f"\nOrder Details:")
    print(f"  Account: *****")
    print(f"  Symbol: MNQU5")
    print(f"  Side: Buy")
    print(f"  Quantity: 1")
    print(f"  Type: Limit")
    print(f"  Price: {limit_price} (${current_bid - limit_price:.2f} below market)")

    # Map columns based on your header row
    # Buy/Sell | Qty | Symbol | Exchange | Market/Limit | Limit Price | Stop Price | Account Id | Duration...
    #    A       B       C         D            E              F            G

    print("\nFilling order fields...")

    # Clear row 2 first (in case there's old data)
    for col in 'ABCDEFGHIJKLMNOPQR':
        place_order_sheet.range(f'{col}2').value = None

    # Fill in all required fields
    place_order_sheet.range('A2').value = 'Buy'  # Buy/Sell
    place_order_sheet.range('B2').value = 1  # Quantity
    place_order_sheet.range('C2').value = 'MNQU5'  # Symbol (adjust if different)
    place_order_sheet.range('D2').value = 'CME'  # Exchange
    place_order_sheet.range('E2').value = 'Limit'  # Order Type
    place_order_sheet.range('F2').value = limit_price  # Limit Price
    place_order_sheet.range('G2').value = None  # Stop Price (empty for limit order)

    # Account and optional fields
    place_order_sheet.range('H2').value = '*****'  # Account ID
    # place_order_sheet.range('I2').value = 'Day'        # Duration (Day, GTC, etc.)

    # Trigger fields (based on your headers)
    place_order_sheet.range('O2').value = 'Yes'  # Place Order Now
    place_order_sheet.range('P2').value = 'Enabled'  # Enable Placing Order

    print("✓ All fields populated")

    # Force recalculation
    wb.app.calculate()

    # Give it a moment to process
    time.sleep(1)

    # Check status
    status = place_order_sheet.range('Q2').value  # Status column
    if status:
        print(f"\nStatus: {status}")

    last_order_num = place_order_sheet.range('R2').value  # Last Order Number
    if last_order_num:
        print(f"Order Number: {last_order_num}")

    print("\n" + "=" * 60)
    print("ORDER SUBMITTED")
    print("=" * 60)
    print("Check R|Trader Pro to verify the order appeared")
    print("If not, try different symbol formats:")
    print("  - MNQ")
    print("  - MNQZ5")
    print("  - MNQ DEC25")
    print("  - @MNQU5")

    return True


def check_symbol_format():
    """Try to determine correct symbol format from existing data"""
    print("\nChecking for symbol format...")

    wb = xw.books.active

    # Check if the Charts sheet title has the symbol
    for sheet in wb.sheets:
        if 'Charts' in sheet.name:
            # Sheet might be named like "Charts-MNQU5"
            if '.' in sheet.name or '-' in sheet.name:
                print(f"Found possible symbol in sheet name: {sheet.name}")
                parts = sheet.name.replace('Charts-', '').replace('Charts_', '')
                if parts:
                    print(f"Extracted symbol: {parts}")
                    return parts

    # Check data sheet for any symbol references
    data_sheet = None
    for sheet in wb.sheets:
        if sheet.name == 'Charts':
            data_sheet = sheet
            break

    if data_sheet:
        # Check first few rows for symbol
        for row in range(1, 10):
            for col in 'ABCDEFGH':
                val = data_sheet.range(f'{col}{row}').value
                if val and isinstance(val, str):
                    if any(x in val.upper() for x in ['MNQ', 'MESU', 'CME', 'S&P']):
                        print(f"Found reference: {val}")

    return None


if __name__ == "__main__":
    print("COMPLETE LIMIT ORDER PLACEMENT")
    print("=" * 60)
    print("\n⚠️  This will place a REAL order!")
    print("⚠️  Make sure you're in DEMO/SIM mode!")

    # Try to detect symbol
    symbol = check_symbol_format()

    print("\nOrder will be placed with:")
    print("  Symbol: MNQU5 (adjust if needed)")
    print("  Buy 1 contract")
    print("  Limit price: $100 below current market")

    response = input("\nProceed? (yes/no): ")

    if response.lower() == 'yes':
        place_complete_limit_order()

        print("\n" + "=" * 60)
        print("TROUBLESHOOTING")
        print("=" * 60)
        print("\nIf order didn't appear, check:")
        print("1. Symbol format - might need MNQ instead of MNQU5")
        print("2. Account ID - might be required in column H")
        print("3. Check column Q for status messages")
        print("4. Check column S for error messages")

        print("\nTo modify, edit these cells in 'Place Orders' sheet:")
        print("  C2: Symbol (try MNQ, MNQZ5, etc.)")
        print("  D2: Exchange (CME)")
        print("  H2: Account ID (if you have one)")
    else:
        print("\nOrder cancelled.")
