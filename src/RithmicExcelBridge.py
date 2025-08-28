import time
from datetime import datetime
from typing import Dict

import pandas as pd
import xlwings as xw


class RithmicExcelBridge:
    """Read real-time data from R|Trader Pro via Excel streaming"""

    def __init__(self, excel_file: str = None):
        """
        Initialize connection to Excel spreadsheet

        Args:
            excel_file: Path to saved Excel file, or None to use active workbook
        """
        self.wb = None
        self.sheet = None
        self.connect(excel_file)

    def connect(self, excel_file: str = None):
        """Connect to Excel workbook"""
        try:
            if excel_file:
                # Open specific file
                self.wb = xw.Book(excel_file)
            else:
                # Use currently active workbook
                self.wb = xw.books.active

            # Get the main sheet (usually first one)
            self.sheet = self.wb.sheets[0]
            print(f"✓ Connected to: {self.wb.name}")

        except Exception as e:
            print(f"✗ Failed to connect to Excel: {e}")
            raise

    def get_quote(self, symbol: str = "MESU5") -> Dict:
        """
        Get current quote data from Excel

        Based on your Excel layout:
        Row 2: Current prices (Bid Size, Bid Price, Ask Size, Ask Price)
        """
        try:
            # Read from your specific layout
            bid_size = self.sheet.range('A2').value
            bid_price = self.sheet.range('B2').value
            ask_size = self.sheet.range('C2').value
            ask_price = self.sheet.range('D2').value
            market_mode = self.sheet.range('E2').value

            # Additional data if available
            last_price = self.sheet.range('D7').value  # From Open row
            high = self.sheet.range('D8').value
            low = self.sheet.range('D9').value
            close = self.sheet.range('D10').value
            volume = self.sheet.range('D12').value

            return {
                'timestamp': datetime.now(),
                'symbol': symbol,
                'bid_size': bid_size,
                'bid': bid_price,
                'ask_size': ask_size,
                'ask': ask_price,
                'last': last_price,
                'high': high,
                'low': low,
                'close': close,
                'volume': volume,
                'spread': ask_price - bid_price if ask_price and bid_price else None,
                'mid': (ask_price + bid_price) / 2 if ask_price and bid_price else None
            }

        except Exception as e:
            print(f"Error reading quote: {e}")
            return None

    def get_bars_data(self) -> pd.DataFrame:
        """Get bar data from Excel (columns G-P)"""
        try:
            # Read the bar data table
            # Assuming it starts at row 6 with headers
            headers = self.sheet.range('G5:P5').value

            # Find last row with data
            last_row = 6
            while self.sheet.range(f'G{last_row}').value is not None:
                last_row += 1
            last_row -= 1

            if last_row >= 6:
                # Read all data
                data = self.sheet.range(f'G6:P{last_row}').value

                # Create DataFrame
                df = pd.DataFrame(data, columns=headers)
                df['Bar Ending Time'] = pd.to_datetime(df['Bar Ending Time'])

                return df

            return pd.DataFrame()

        except Exception as e:
            print(f"Error reading bars: {e}")
            return pd.DataFrame()

    def stream_quotes(self, callback=None, interval=0.1):
        """
        Stream quotes in real-time

        Args:
            callback: Function to call with each quote
            interval: Seconds between reads
        """
        print("Starting quote stream... (Press Ctrl+C to stop)")

        try:
            last_bid = None
            last_ask = None
            update_count = 0

            while True:
                quote = self.get_quote()

                if quote:
                    # Check if prices changed (not the entire quote dict)
                    if quote['bid'] != last_bid or quote['ask'] != last_ask:
                        update_count += 1

                        if callback:
                            callback(quote)
                        else:
                            # Default: print the quote with update count
                            print(f"\n[{quote['timestamp'].strftime('%H:%M:%S')}] Update #{update_count} | "
                                  f"Bid: {quote['bid']:.2f} | "
                                  f"Ask: {quote['ask']:.2f} | "
                                  f"Spread: {quote['spread']:.2f}")

                        last_bid = quote['bid']
                        last_ask = quote['ask']

                time.sleep(interval)

        except KeyboardInterrupt:
            print(f"\n✓ Stream stopped. Total updates: {update_count}")

    def save_to_csv(self, filename: str, duration: int = 60):
        """
        Save streaming data to CSV for specified duration

        Args:
            filename: Output CSV filename
            duration: Seconds to record
        """
        print(f"Recording data for {duration} seconds...")

        quotes = []
        start_time = time.time()

        while time.time() - start_time < duration:
            quote = self.get_quote()
            if quote:
                quotes.append(quote)
            time.sleep(0.1)

        # Convert to DataFrame and save
        df = pd.DataFrame(quotes)
        df.to_csv(filename, index=False)
        print(f"✓ Saved {len(quotes)} quotes to {filename}")

        return df


class RithmicTrader:
    """Simple trading logic using Excel bridge"""

    def __init__(self, excel_bridge: RithmicExcelBridge):
        self.bridge = excel_bridge
        self.position = 0
        self.entry_price = None

    def check_signal(self, quote: Dict) -> str:
        """
        Example trading signal
        Returns: 'BUY', 'SELL', or 'HOLD'
        """
        if not quote['bid'] or not quote['ask']:
            return 'HOLD'

        spread = quote['spread']

        # Example logic - replace with your strategy
        if spread < 0.25:  # Tight spread
            if self.position == 0:
                return 'BUY'
            elif self.position > 0 and quote['bid'] > self.entry_price + 1.0:
                return 'SELL'  # Take profit

        return 'HOLD'

    def on_quote(self, quote: Dict):
        """Handle incoming quote"""
        signal = self.check_signal(quote)

        if signal == 'BUY' and self.position == 0:
            self.position = 1
            self.entry_price = quote['ask']
            print(f"\n🟢 BUY at {quote['ask']:.2f}")

        elif signal == 'SELL' and self.position > 0:
            pnl = quote['bid'] - self.entry_price
            print(f"\n🔴 SELL at {quote['bid']:.2f} | PnL: {pnl:.2f}")
            self.position = 0
            self.entry_price = None


def main():
    """Example usage"""

    # Setup instructions
    print("=" * 60)
    print("RITHMIC EXCEL BRIDGE SETUP")
    print("=" * 60)
    print("\n1. In R|Trader Pro:")
    print("   - Open your chart (MESU5.CME)")
    print("   - Right-click → Create Live Streaming Spreadsheet")
    print("   - Save the Excel file")
    print("\n2. Keep Excel open with the streaming data")
    print("\n3. Run this script")
    print("-" * 60)

    # Connect to Excel
    bridge = RithmicExcelBridge()  # Uses active workbook

    # Get single quote
    quote = bridge.get_quote()
    if quote:
        print(f"\nCurrent Quote:")
        print(f"  Bid: {quote['bid']} ({quote['bid_size']})")
        print(f"  Ask: {quote['ask']} ({quote['ask_size']})")
        print(f"  Spread: {quote['spread']:.2f}")

    # Stream quotes
    print("\n" + "=" * 60)
    print("STREAMING QUOTES")
    print("=" * 60)

    # Option 1: Simple streaming
    # bridge.stream_quotes(interval=0.5)

    # Option 2: With trading logic
    trader = RithmicTrader(bridge)
    bridge.stream_quotes(callback=trader.on_quote, interval=0.5)

    # Option 3: Save to CSV
    # df = bridge.save_to_csv("rithmic_data.csv", duration=60)


if __name__ == "__main__":
    main()
