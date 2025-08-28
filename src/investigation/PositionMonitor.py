import xlwings as xw
import time
from datetime import datetime

class PositionMonitor:
    """Monitor positions and orders, auto-place stop losses"""

    def __init__(self):
        self.wb = xw.books.active
        self.positions = {}
        self.orders = {}
        self.managed_positions = set()  # Track which positions have SL

        # Find relevant sheets
        self.sheets = {}
        for sheet in self.wb.sheets:
            if 'Manage Orders' in sheet.name:
                self.sheets['manage'] = sheet
            elif 'Place Orders' in sheet.name:
                self.sheets['place'] = sheet
            elif sheet.name == 'Charts':
                self.sheets['data'] = sheet
            # Look for position/account sheets
            elif any(word in sheet.name.lower() for word in ['position', 'account', 'portfolio']):
                self.sheets['positions'] = sheet
                print(f"Found position sheet: {sheet.name}")

    def scan_for_position_data(self):
        """Scan all sheets for position and order data"""
        print("\n" + "="*60)
        print("SCANNING FOR POSITION AND ORDER DATA")
        print("="*60)

        position_data = {}

        for sheet in self.wb.sheets:
            print(f"\nScanning: {sheet.name}")
            print("-"*40)

            # Look for position/order keywords in first 20 rows
            for row in range(1, 21):
                row_data = []
                for col in range(1, 16):  # Check columns A-O
                    try:
                        value = sheet.range((row, col)).value
                        if value:
                            value_str = str(value).lower()

                            # Position indicators
                            if any(word in value_str for word in ['position', 'quantity', 'avg price', 'p&l', 'entry']):
                                print(f"  Position data at {chr(64+col)}{row}: {value}")

                            # Order book indicators
                            elif any(word in value_str for word in ['working', 'filled', 'order id', 'status']):
                                print(f"  Order data at {chr(64+col)}{row}: {value}")

                            # Account data
                            elif any(word in value_str for word in ['balance', 'margin', 'buying power']):
                                print(f"  Account data at {chr(64+col)}{row}: {value}")

                            row_data.append(value)
                    except:
                        pass

                # Print full rows that might be position data
                if row_data and row > 1:  # Skip headers
                    # Check if this looks like position data (has numbers)
                    numbers_in_row = sum(1 for v in row_data if isinstance(v, (int, float)))
                    if numbers_in_row >= 2:  # At least 2 numeric values
                        print(f"  Data row {row}: {row_data[:8]}")  # First 8 columns

        return position_data

    def get_current_position(self):
        """Try to read current position from Excel"""
        # Common places for position data
        possible_locations = [
            ('Charts', 'K10'),  # Random guess
            ('Charts', 'L10'),
            ('Manage Orders-Charts', 'D2'),
            ('Manage Orders-Charts', 'E2'),
        ]

        for sheet_name, cell_ref in possible_locations:
            try:
                sheet = self.wb.sheets[sheet_name]
                value = sheet.range(cell_ref).value
                if value and isinstance(value, (int, float)) and value != 0:
                    print(f"Found position at {sheet_name}!{cell_ref}: {value}")
                    return value
            except:
                pass

        return 0

    def monitor_for_new_positions(self):
        """Monitor for new positions and auto-place stop losses"""
        print("\n" + "="*60)
        print("POSITION MONITOR WITH AUTO STOP-LOSS")
        print("="*60)

        print("\nMonitoring for position changes...")
        print("When position detected, will auto-place stop loss")
        print("Press Ctrl+C to stop\n")

        last_position = 0
        check_count = 0

        try:
            while True:
                check_count += 1

                # Get current position (you'll need to find where this is)
                current_position = self.get_current_position()

                # Check if position changed
                if current_position != last_position:
                    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Position change detected!")
                    print(f"  Previous: {last_position}")
                    print(f"  Current: {current_position}")

                    # New position opened
                    if last_position == 0 and current_position != 0:
                        self.handle_new_position(current_position)

                    # Position closed
                    elif current_position == 0 and last_position != 0:
                        print("  Position closed")
                        # Could cancel pending SL orders here

                    # Position reversed
                    elif current_position * last_position < 0:
                        print("  Position reversed!")
                        self.handle_new_position(current_position)

                    last_position = current_position

                # Periodic status
                if check_count % 30 == 0:  # Every 30 checks
                    print(f".", end="", flush=True)

                time.sleep(0.5)  # Check every 500ms

        except KeyboardInterrupt:
            print("\nMonitoring stopped")

    def handle_new_position(self, position_size):
        """Auto-place stop loss for new position"""
        print(f"\n  AUTO-PLACING STOP LOSS for position: {position_size}")

        if not self.sheets.get('place'):
            print("  ✗ Place Orders sheet not found")
            return

        # Get current price for stop calculation
        if self.sheets.get('data'):
            bid = self.sheets['data'].range('B2').value
            ask = self.sheets['data'].range('D2').value
            current_price = (bid + ask) / 2
        else:
            current_price = 23500  # Default

        # Calculate stop loss
        stop_distance = 10  # Points for stop loss

        if position_size > 0:  # Long position
            stop_price = current_price - stop_distance
            side = 'Sell'
        else:  # Short position
            stop_price = current_price + stop_distance
            side = 'Buy'

        stop_price = round(stop_price, 2)

        print(f"  Placing {side} Stop at {stop_price}")
        print(f"  Stop is {stop_distance} points from {current_price:.2f}")

        # Place the stop order
        place_sheet = self.sheets['place']

        # Clear and fill order row
        place_sheet.range('A2').value = side
        place_sheet.range('B2').value = abs(position_size)
        place_sheet.range('C2').value = 'MESU5.CME'  # Or MNQ
        place_sheet.range('D2').value = 'CME'
        place_sheet.range('E2').value = 'Stop Market'
        place_sheet.range('F2').value = None  # No limit price for stop market
        place_sheet.range('G2').value = stop_price
        place_sheet.range('H2').value = '209324'  # Account ID
        place_sheet.range('O2').value = 'Yes'
        place_sheet.range('P2').value = 'Enabled'

        # Force calculation
        self.wb.app.calculate()

        print(f"  ✓ Stop loss order placed")


def test_position_detection():
    """Test if we can detect positions"""
    monitor = PositionMonitor()

    # First scan to find where position data is
    monitor.scan_for_position_data()

    print("\n" + "="*60)
    print("POSITION DETECTION TEST")
    print("="*60)

    print("\n1. Open a position in ATAS or R|Trader Pro")
    print("2. Watch if it appears in Excel")
    print("3. The script will auto-place a stop loss\n")

    response = input("Start monitoring? (yes/no): ")

    if response.lower() == 'yes':
        monitor.monitor_for_new_positions()


def create_position_sheet():
    """Create a custom sheet to track positions if needed"""
    print("\nIf R|Trader Pro doesn't export positions, we could:")
    print("1. Manually track when orders fill")
    print("2. Create our own position tracker")
    print("3. Use the 'Manage Orders' sheet to see filled orders")

    wb = xw.books.active

    # Check if we can see filled orders
    for sheet in wb.sheets:
        if 'Manage' in sheet.name:
            print(f"\nChecking {sheet.name} for order status...")

            # Look for filled orders
            for row in range(2, 10):
                status = sheet.range(f'S{row}').value  # Status column
                if status:
                    print(f"  Row {row} status: {status}")


if __name__ == "__main__":
    print("POSITION MONITOR WITH AUTO STOP-LOSS")
    print("="*60)

    print("\nThis script will:")
    print("1. Scan Excel for position data")
    print("2. Monitor for new positions")
    print("3. Auto-place stop losses")

    test_position_detection()

    print("\n" + "="*60)
    print("NOTES")
    print("="*60)
    print("""
If positions aren't visible in Excel:
- Check if R|Trader Pro has 'Export Positions' option
- Create another streaming spreadsheet for positions
- Look in View → Account Info → Export to Excel
- The 'Manage Orders' sheet might show filled quantities

For now, you could:
1. Monitor the order sheet for 'Filled' status
2. When order fills, auto-place the SL
3. Track positions manually based on filled orders
""")