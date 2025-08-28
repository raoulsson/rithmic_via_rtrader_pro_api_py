import csv
import time
from datetime import datetime

import xlwings as xw


def stream_prices():
    """Simple direct streaming - exactly what worked in the debugger"""

    # Connect to Excel
    wb = xw.books.active
    sheet = wb.sheets[0]
    print(f"Connected to: {wb.name}")
    print("Streaming prices... (Press Ctrl+C to stop)\n")

    last_bid = None
    last_ask = None
    update_count = 0

    try:
        while True:
            # Read directly from cells - exactly what worked in debugger
            bid = sheet.range('B2').value
            ask = sheet.range('D2').value

            # Check if values changed
            if bid != last_bid or ask != last_ask:
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

            time.sleep(0.1)  # Check 10 times per second

    except KeyboardInterrupt:
        print(f"\nStopped. Total updates: {update_count}")


def stream_with_data():
    """Stream with additional market data"""

    wb = xw.books.active
    sheet = wb.sheets[0]
    print(f"Connected to: {wb.name}")
    print("Streaming full market data... (Press Ctrl+C to stop)\n")

    last_bid = None
    last_ask = None
    update_count = 0

    try:
        while True:
            # Read all relevant cells
            bid_size = sheet.range('A2').value
            bid = sheet.range('B2').value
            ask_size = sheet.range('C2').value
            ask = sheet.range('D2').value

            # Additional data from lower rows
            last_price = sheet.range('D6').value  # Open price
            high = sheet.range('D7').value
            low = sheet.range('D8').value

            # Check if main prices changed
            if bid != last_bid or ask != last_ask:
                update_count += 1
                timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]

                # Calculate spread and mid
                spread = ask - bid if ask and bid else 0
                mid = (ask + bid) / 2 if ask and bid else 0

                print(f"[{timestamp}] #{update_count}")
                print(f"  Bid: {bid:.2f} ({int(bid_size) if bid_size else 0})")
                print(f"  Ask: {ask:.2f} ({int(ask_size) if ask_size else 0})")
                print(f"  Mid: {mid:.2f} | Spread: {spread:.2f}")
                print(f"  Range: {low:.2f} - {high:.2f}")
                print()

                last_bid = bid
                last_ask = ask

            time.sleep(0.1)

    except KeyboardInterrupt:
        print(f"\nStopped. Total updates: {update_count}")


def test_connection():
    """Test if we can read from Excel at all"""
    print("Testing Excel connection...")

    try:
        wb = xw.books.active
        sheet = wb.sheets[0]
        print(f"✓ Connected to: {wb.name}")
        print(f"✓ Sheet name: {sheet.name}")

        # Read current values
        bid = sheet.range('B2').value
        ask = sheet.range('D2').value

        print(f"✓ Current Bid: {bid}")
        print(f"✓ Current Ask: {ask}")

        if bid and ask:
            print("\n✓ Excel connection working!")
            return True
        else:
            print("\n✗ No data in cells")
            return False

    except Exception as e:
        print(f"\n✗ Error: {e}")
        return False


if __name__ == "__main__":
    print("=" * 60)
    print("SIMPLE EXCEL PRICE STREAMER")
    print("=" * 60)
    print("\nMake sure:")
    print("1. R|Trader Pro is running with a chart open")
    print("2. Live Streaming Spreadsheet is created and open in Excel")
    print("3. Streaming is NOT paused in R|Trader Pro")
    print("-" * 60)

    # Test connection first
    if test_connection():
        print("\nChoose streaming mode:")
        print("1. Simple price updates only")
        print("2. Full market data")

        choice = input("\nEnter choice (1 or 2): ").strip()

        if choice == "2":
            stream_with_data()
        else:
            stream_prices()
    else:
        print("\nPlease fix the connection and try again")


def stream_and_save(filename='price_data.csv'):
    """Stream prices and save to CSV"""

    wb = xw.books.active
    sheet = wb.sheets[0]

    with open(filename, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['timestamp', 'bid', 'ask', 'spread', 'mid'])

        last_bid = None
        last_ask = None

        while True:
            bid = sheet.range('B2').value
            ask = sheet.range('D2').value

            if bid != last_bid or ask != last_ask:
                timestamp = datetime.now()
                spread = ask - bid
                mid = (ask + bid) / 2

                writer.writerow([timestamp, bid, ask, spread, mid])
                f.flush()  # Write immediately

                print(f"Saved: {timestamp.strftime('%H:%M:%S')} Bid:{bid} Ask:{ask}")

                last_bid = bid
                last_ask = ask

            time.sleep(0.1)
