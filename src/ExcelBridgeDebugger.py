import time
from datetime import datetime

import xlwings as xw


def debug_excel_updates():
    """Debug why Excel updates aren't showing"""

    print("EXCEL UPDATE DEBUGGER")
    print("=" * 60)

    # Connect to Excel
    wb = xw.books.active
    sheet = wb.sheets[0]
    print(f"Connected to: {wb.name}")
    print(f"Sheet: {sheet.name}")

    # Test 1: Are Excel cells actually updating?
    print("\n1. Testing if Excel cells are updating...")
    print("-" * 40)

    bid_values = []
    ask_values = []

    for i in range(10):
        bid = sheet.range('B2').value
        ask = sheet.range('D2').value
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]

        bid_values.append(bid)
        ask_values.append(ask)

        print(f"{timestamp} | Bid: {bid} | Ask: {ask}")
        time.sleep(0.5)

    # Check if values changed
    bid_changes = len(set(bid_values))
    ask_changes = len(set(ask_values))

    print(f"\nUnique bid values seen: {bid_changes}")
    print(f"Unique ask values seen: {ask_changes}")

    if bid_changes == 1 and ask_changes == 1:
        print("❌ Values are NOT updating in Excel!")
        print("\nPossible fixes:")
        print("1. Make sure 'Pause Streaming' is NOT checked in R|Trader Pro")
        print("2. Try 'Set Stream Name' in R|Trader Pro to restart streaming")
        print("3. Close and recreate the Live Streaming Spreadsheet")
        print("4. Check if market is open (might be no updates outside market hours)")
    else:
        print("✓ Values ARE updating in Excel")

    # Test 2: Check different cell locations
    print("\n2. Checking all cells for activity...")
    print("-" * 40)

    # Read a broader range to see what's updating
    for row in range(1, 11):
        row_data = sheet.range(f'A{row}:E{row}').value
        if any(row_data):
            print(f"Row {row}: {row_data}")

    # Test 3: Monitor with visual feedback
    print("\n3. Live monitoring (watch for changes)...")
    print("-" * 40)
    print("Press Ctrl+C to stop\n")

    last_bid = None
    last_ask = None
    update_count = 0
    no_change_count = 0

    try:
        while True:
            bid = sheet.range('B2').value
            ask = sheet.range('D2').value

            changed = False
            if bid != last_bid or ask != last_ask:
                changed = True
                update_count += 1
                timestamp = datetime.now().strftime('%H:%M:%S')

                # Show what changed
                changes = []
                if bid != last_bid:
                    changes.append(f"Bid: {last_bid} → {bid}")
                if ask != last_ask:
                    changes.append(f"Ask: {last_ask} → {ask}")

                print(f"[{timestamp}] UPDATE #{update_count}: {', '.join(changes)}")

                last_bid = bid
                last_ask = ask
            else:
                no_change_count += 1
                # Print a dot every 10 checks to show it's running
                if no_change_count % 10 == 0:
                    print(".", end="", flush=True)

            time.sleep(0.1)

    except KeyboardInterrupt:
        print(f"\n\nSummary: {update_count} updates detected")


def test_linkable_mode():
    """Test if Linkable Mode provides different behavior"""
    print("\n" + "=" * 60)
    print("LINKABLE MODE TEST")
    print("=" * 60)
    print("\nIf basic streaming isn't updating, try:")
    print("1. In R|Trader Pro: Create Live Streaming Spreadsheet (Linkable Mode)")
    print("2. This might provide DDE links or RTD formulas")
    print("\nCheck if cells now contain formulas instead of values")

    wb = xw.books.active
    sheet = wb.sheets[0]

    # Check if B2 has a formula
    cell_b2 = sheet.range('B2')
    if cell_b2.formula and cell_b2.formula.startswith('='):
        print(f"\n✓ Cell B2 has formula: {cell_b2.formula}")
        print("This is Linkable Mode - formulas update automatically")
    else:
        print(f"\n✗ Cell B2 has value: {cell_b2.value}")
        print("This is regular streaming mode")


def test_alternative_reading():
    """Try alternative methods to read Excel data"""
    print("\n" + "=" * 60)
    print("ALTERNATIVE READING METHODS")
    print("=" * 60)

    wb = xw.books.active
    sheet = wb.sheets[0]

    print("\n1. Force refresh before reading:")
    for i in range(5):
        # Force Excel to recalculate
        wb.app.calculate()

        # Some versions need this
        sheet.api.Calculate()

        # Now read
        bid = sheet.range('B2').value
        ask = sheet.range('D2').value
        print(f"  Read {i + 1}: Bid={bid}, Ask={ask}")
        time.sleep(0.5)

    print("\n2. Using Excel COM API directly:")
    try:
        # Get the COM object
        excel_app = wb.app.api
        excel_sheet = sheet.api

        for i in range(5):
            bid = excel_sheet.Range("B2").Value
            ask = excel_sheet.Range("D2").Value
            print(f"  COM Read {i + 1}: Bid={bid}, Ask={ask}")
            time.sleep(0.5)
    except Exception as e:
        print(f"  COM Error: {e}")


if __name__ == "__main__":
    # Run all debugging tests
    debug_excel_updates()
    test_linkable_mode()
    test_alternative_reading()

    print("\n" + "=" * 60)
    print("RECOMMENDATIONS")
    print("=" * 60)
    print("""
If data isn't updating:

1. CHECK R|TRADER PRO:
   - Is 'Pause Streaming' unchecked?
   - Is the chart still active?
   - Try 'Set Stream Name' to restart

2. CHECK EXCEL:
   - Are cells updating visually in Excel?
   - Try manual refresh (F9)
   - Close and reopen the spreadsheet

3. TRY LINKABLE MODE:
   - Create Live Streaming Spreadsheet (Linkable Mode)
   - This might work better for programmatic access

4. MARKET HOURS:
   - Is the market open?
   - E-mini S&P futures trade Sun 6pm - Fri 5pm ET
   - Limited updates outside regular hours
""")

    input("\nPress Enter to exit...")
