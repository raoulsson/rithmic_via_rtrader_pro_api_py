import time

import xlwings as xw


class RithmicExcelTrader:
    """Execute orders through R|Trader Pro's Excel interface"""

    def __init__(self):
        self.wb = xw.books.active
        self.data_sheet = None
        self.place_order_sheet = None
        self.manage_order_sheet = None

        # Find the sheets
        for sheet in self.wb.sheets:
            if 'Place Orders' in sheet.name:
                self.place_order_sheet = sheet
                print(f"✓ Found order placement sheet: {sheet.name}")
            elif 'Manage Orders' in sheet.name:
                self.manage_order_sheet = sheet
                print(f"✓ Found order management sheet: {sheet.name}")
            elif sheet.name == 'Charts':
                self.data_sheet = sheet
                print(f"✓ Found data sheet: {sheet.name}")

    def check_order_sheet_structure(self):
        """Examine the order sheet structure"""
        print("\n" + "=" * 60)
        print("ORDER SHEET STRUCTURE")
        print("=" * 60)

        if self.place_order_sheet:
            print(f"\n{self.place_order_sheet.name}:")
            print("-" * 40)

            # Read first 3 rows and 10 columns to understand structure
            for row in range(1, 4):
                row_data = []
                for col in range(1, 11):
                    value = self.place_order_sheet.range((row, col)).value
                    row_data.append(str(value) if value else "")
                print(f"Row {row}: {row_data}")

            # Check specific cells mentioned in scan
            print("\nKey cells:")
            print(f"A1 (Buy/Sell): {self.place_order_sheet.range('A1').value}")
            print(f"A2: {self.place_order_sheet.range('A2').value}")
            print(f"E1 (Order Type): {self.place_order_sheet.range('E1').value}")
            print(f"E2: {self.place_order_sheet.range('E2').value}")
            print(f"F1 (Limit Price): {self.place_order_sheet.range('F1').value}")
            print(f"F2: {self.place_order_sheet.range('F2').value}")
            print(f"G1 (Stop Price): {self.place_order_sheet.range('G1').value}")
            print(f"G2: {self.place_order_sheet.range('G2').value}")

        if self.manage_order_sheet:
            print(f"\n{self.manage_order_sheet.name}:")
            print("-" * 40)

            for row in range(1, 4):
                row_data = []
                for col in range(1, 11):
                    value = self.manage_order_sheet.range((row, col)).value
                    row_data.append(str(value) if value else "")
                print(f"Row {row}: {row_data}")

    def place_market_order(self, side='Buy', quantity=1):
        """Place a market order"""
        if not self.place_order_sheet:
            print("✗ Place Orders sheet not found")
            return False

        print(f"\n📝 Placing {side} Market Order for {quantity} contracts...")

        try:
            # Set order parameters
            # Assuming row 2 is for entering orders (row 1 is headers)
            self.place_order_sheet.range('A2').value = side  # Buy or Sell
            self.place_order_sheet.range('B2').value = quantity  # Quantity
            self.place_order_sheet.range('E2').value = 'Market'  # Order type

            print(f"✓ Order parameters set: {side} {quantity} @ Market")

            # There might be a trigger cell to submit
            # Check if there's a "Submit" or "Place Order" cell
            for row in range(1, 10):
                for col in range(1, 10):
                    cell_value = self.place_order_sheet.range((row, col)).value
                    if cell_value and 'submit' in str(cell_value).lower():
                        print(f"Found submit trigger at {chr(64 + col)}{row}")
                        # Set it to trigger
                        self.place_order_sheet.range((row, col + 1)).value = 'Yes'
                        break

            return True

        except Exception as e:
            print(f"✗ Error placing order: {e}")
            return False

    def place_limit_order(self, side='Buy', quantity=1, limit_price=None):
        """Place a limit order"""
        if not self.place_order_sheet:
            print("✗ Place Orders sheet not found")
            return False

        # Get current market price if limit not specified
        if limit_price is None and self.data_sheet:
            bid = self.data_sheet.range('B2').value
            ask = self.data_sheet.range('D2').value
            limit_price = bid if side == 'Buy' else ask

        print(f"\n📝 Placing {side} Limit Order: {quantity} @ {limit_price}")

        try:
            # Set order parameters
            self.place_order_sheet.range('A2').value = side  # Buy or Sell
            self.place_order_sheet.range('B2').value = quantity  # Quantity (might be in B2)
            self.place_order_sheet.range('E2').value = 'Limit'  # Order type
            self.place_order_sheet.range('F2').value = limit_price  # Limit price

            print(f"✓ Limit order parameters set")
            return True

        except Exception as e:
            print(f"✗ Error placing limit order: {e}")
            return False

    def place_stop_order(self, side='Buy', quantity=1, stop_price=None):
        """Place a stop order"""
        if not self.place_order_sheet:
            return False

        print(f"\n📝 Placing {side} Stop Order: {quantity} @ {stop_price}")

        try:
            self.place_order_sheet.range('A2').value = side
            self.place_order_sheet.range('B2').value = quantity
            self.place_order_sheet.range('E2').value = 'Stop Market'
            self.place_order_sheet.range('G2').value = stop_price  # Stop price

            print(f"✓ Stop order parameters set")
            return True

        except Exception as e:
            print(f"✗ Error: {e}")
            return False

    def cancel_order(self, order_number):
        """Cancel an order"""
        if not self.manage_order_sheet:
            print("✗ Manage Orders sheet not found")
            return False

        print(f"\n🚫 Cancelling order {order_number}...")

        try:
            # Find the cancel order section
            self.manage_order_sheet.range('A2').value = order_number
            self.manage_order_sheet.range('B2').value = 'Yes'  # Trigger cancel

            print(f"✓ Cancel request sent for order {order_number}")
            return True

        except Exception as e:
            print(f"✗ Error cancelling order: {e}")
            return False

    def get_current_prices(self):
        """Get current bid/ask from data sheet"""
        if self.data_sheet:
            bid = self.data_sheet.range('B2').value
            ask = self.data_sheet.range('D2').value
            return {'bid': bid, 'ask': ask}
        return None


def test_order_execution():
    """Test if order execution works"""
    print("=" * 60)
    print("TESTING ORDER EXECUTION")
    print("=" * 60)

    trader = RithmicExcelTrader()

    # First understand the sheet structure
    trader.check_order_sheet_structure()

    print("\n" + "=" * 60)
    print("TEST ORDER PLACEMENT")
    print("=" * 60)
    print("\n⚠️  WARNING: This might place a REAL order!")
    print("Make sure you're in DEMO/SIM mode first!\n")

    response = input("Do you want to test order placement? (yes/no): ")

    if response.lower() == 'yes':
        print("\nChoose test:")
        print("1. Check structure only (safe)")
        print("2. Place a far limit buy order (safer)")
        print("3. Place a market order (risky!)")

        choice = input("Choice (1-3): ")

        if choice == '2':
            # Get current price
            prices = trader.get_current_prices()
            if prices:
                # Place buy limit far below market
                safe_price = prices['bid'] - 100  # 100 points below
                trader.place_limit_order('Buy', 1, safe_price)
                print(f"\nPlaced Buy Limit at {safe_price} (far from market {prices['bid']})")
                print("Check R|Trader Pro to see if order appeared!")

                time.sleep(3)

                # Try to cancel it
                order_num = input("Enter order number to cancel (or skip): ")
                if order_num:
                    trader.cancel_order(order_num)

        elif choice == '3':
            print("\n⚠️  MARKET ORDER - ARE YOU SURE?")
            confirm = input("Type 'CONFIRM' to proceed: ")
            if confirm == 'CONFIRM':
                trader.place_market_order('Buy', 1)
                print("Check R|Trader Pro for order execution!")


def monitor_and_trade():
    """Simple trading bot using Excel interface"""
    print("=" * 60)
    print("AUTOMATED TRADING BOT")
    print("=" * 60)

    trader = RithmicExcelTrader()

    print("\nBot will:")
    print("1. Monitor prices")
    print("2. Place orders based on signals")
    print("3. THIS IS LIVE TRADING - USE DEMO ACCOUNT!\n")

    position = 0
    entry_price = None

    try:
        while True:
            prices = trader.get_current_prices()
            if prices:
                spread = prices['ask'] - prices['bid']

                # Simple strategy: buy on tight spread
                if position == 0 and spread <= 0.25:
                    print(f"\n🟢 Signal: Tight spread {spread:.2f}")
                    trader.place_limit_order('Buy', 1, prices['bid'])
                    position = 1
                    entry_price = prices['bid']

                elif position == 1 and prices['bid'] > entry_price + 2:
                    print(f"\n🔴 Take profit signal")
                    trader.place_limit_order('Sell', 1, prices['ask'])
                    position = 0

                print(f"\rBid: {prices['bid']:.2f} Ask: {prices['ask']:.2f} Pos: {position}", end='')

            time.sleep(1)

    except KeyboardInterrupt:
        print("\nBot stopped")
        if position != 0:
            print("⚠️  Warning: Still have open position!")


if __name__ == "__main__":
    print("R|TRADER PRO EXCEL ORDER EXECUTION\n")

    # Test the order interface
    test_order_execution()

    print("\n" + "=" * 60)
    print("NEXT STEPS")
    print("=" * 60)
    print("""
1. Run this to understand the sheet structure
2. Manually test in Excel first:
   - Put 'Buy' in A2 of Place Orders sheet
   - Put '1' in B2 (quantity)
   - Put 'Market' in E2
   - See if anything happens
   
3. Look for trigger cells like:
   - 'Submit Order'
   - 'Place Order Now'
   - 'Execute'
   
4. The order might need:
   - A specific trigger value
   - Pressing Enter
   - Excel calculation (F9)
   
5. TEST IN DEMO/SIM MODE FIRST!
""")

    input("\nPress Enter to exit...")
