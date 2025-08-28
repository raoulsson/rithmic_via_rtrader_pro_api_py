import time

import win32com.client
import xlwings as xw


def check_excel_macros():
    """Check if the Excel file has any order-related macros or buttons"""
    print("=" * 60)
    print("CHECKING FOR ORDER EXECUTION IN EXCEL")
    print("=" * 60)

    wb = xw.books.active
    sheet = wb.sheets[0]

    print(f"\n1. Checking workbook: {wb.name}")
    print(f"   Sheet: {sheet.name}")

    # Check for VBA macros
    print("\n2. Checking for VBA macros...")
    try:
        vba_project = wb.api.VBProject
        print(f"   VBA Project Name: {vba_project.Name}")

        # List all VBA components
        for component in vba_project.VBComponents:
            print(f"   - {component.Name} ({component.Type})")

            # Try to read code modules
            if component.CodeModule.CountOfLines > 0:
                print(f"     Has {component.CodeModule.CountOfLines} lines of code")

                # Look for order-related keywords in first 50 lines
                code_preview = component.CodeModule.Lines(1, min(50, component.CodeModule.CountOfLines))
                order_keywords = ['order', 'buy', 'sell', 'limit', 'market', 'stop', 'execute', 'submit']
                for keyword in order_keywords:
                    if keyword.lower() in code_preview.lower():
                        print(f"     Found '{keyword}' in code!")
    except Exception as e:
        print(f"   No VBA macros found or access denied: {e}")

    # Check for buttons/shapes that might trigger orders
    print("\n3. Checking for buttons or controls...")
    try:
        shapes = sheet.api.Shapes
        if shapes.Count > 0:
            print(f"   Found {shapes.Count} shapes/buttons:")
            for shape in shapes:
                print(
                    f"   - {shape.Name}: {shape.TextFrame2.TextRange.Text if hasattr(shape, 'TextFrame2') else 'No text'}")
        else:
            print("   No buttons or shapes found")
    except Exception as e:
        print(f"   Could not check shapes: {e}")

    # Check for named ranges that might be for orders
    print("\n4. Checking for named ranges...")
    try:
        for name in wb.names:
            if any(word in name.name.lower() for word in ['order', 'buy', 'sell', 'quantity', 'price']):
                print(f"   Found order-related range: {name.name} = {name.refers_to}")
    except:
        print("   No relevant named ranges found")


def test_rtrader_com_interface():
    """Try to find R|Trader Pro's COM interface for orders"""
    print("\n" + "=" * 60)
    print("TESTING R|TRADER PRO COM INTERFACE")
    print("=" * 60)

    # Common ProgIDs to try
    progids = [
        'RithmicTrader.Application',
        'Rithmic.Application',
        'RTrader.Application',
        'RithmicTraderPro.Application',
        'Rithmic.OrderManager',
        'Rithmic.Trading',
    ]

    print("\nTrying to connect to R|Trader Pro via COM...")
    for progid in progids:
        try:
            app = win32com.client.Dispatch(progid)
            print(f"✓ Connected to {progid}")

            # Try to list available methods
            print("  Available methods/properties:")
            for attr in dir(app):
                if not attr.startswith('_'):
                    print(f"    - {attr}")

            # Look for order-related methods
            order_methods = ['SubmitOrder', 'PlaceOrder', 'Buy', 'Sell', 'CreateOrder']
            for method in order_methods:
                if hasattr(app, method):
                    print(f"  ✓ Found method: {method}")

            return app

        except Exception as e:
            continue

    print("✗ Could not find R|Trader Pro COM interface")
    print("  Order execution likely requires the official API")
    return None


def check_excel_formulas_for_orders():
    """Check if there are any cells for entering orders"""
    print("\n" + "=" * 60)
    print("CHECKING FOR ORDER ENTRY CELLS")
    print("=" * 60)

    wb = xw.books.active

    # Check all sheets
    print("\nChecking all sheets for order-related content...")
    for sheet in wb.sheets:
        print(f"\nSheet: {sheet.name}")

        # Check first 20 rows and 10 columns for order keywords
        for row in range(1, 21):
            for col in range(1, 11):
                try:
                    cell_value = sheet.range((row, col)).value
                    if cell_value and isinstance(cell_value, str):
                        order_words = ['order', 'buy', 'sell', 'quantity', 'limit', 'market', 'stop', 'submit']
                        for word in order_words:
                            if word.lower() in str(cell_value).lower():
                                print(f"  Cell {chr(64 + col)}{row}: {cell_value}")

                                # Check if next cells might be input fields
                                next_cell = sheet.range((row, col + 1)).value
                                if next_cell is None or next_cell == '':
                                    print(f"    → Possible input cell at {chr(65 + col)}{row}")
                except:
                    pass


def simulate_order_entry():
    """Show how orders COULD work if the interface existed"""
    print("\n" + "=" * 60)
    print("ORDER EXECUTION POSSIBILITIES")
    print("=" * 60)

    print("""
Based on the investigation, here's what's likely:

1. EXCEL STREAMING IS READ-ONLY:
   - The Live Streaming Spreadsheet is for data export only
   - No built-in order execution via Excel

2. TO EXECUTE ORDERS YOU NEED:
   
   Option A: Manual execution in R|Trader Pro
   - Python generates signals → You manually place orders
   
   Option B: Official Rithmic API (when you get credentials)
   - Direct order submission via API
   - Full automation possible
   
   Option C: UI Automation (risky)
   - Use pyautogui/pywinauto to control R|Trader Pro
   - Fragile and not recommended for real trading
   
   Option D: Check if R|Trader Pro has:
   - File → Import Orders (from CSV)
   - Tools → Order Macro
   - Any plugin system for orders

3. WHAT YOU CAN DO NOW:
""")

    # Show a mock order system
    print("   Create a signal file that you execute manually:")
    print()
    print("   # signals.csv")
    print("   timestamp,action,quantity,order_type,price")
    print(f"   {time.strftime('%Y-%m-%d %H:%M:%S')},BUY,1,LIMIT,23546.50")
    print(f"   {time.strftime('%Y-%m-%d %H:%M:%S')},SELL,1,MARKET,")


def check_rtrader_menus():
    """Document what to check in R|Trader Pro menus"""
    print("\n" + "=" * 60)
    print("CHECK THESE IN R|TRADER PRO")
    print("=" * 60)

    print("""
Look for these menu items in R|Trader Pro:

1. FILE MENU:
   □ Import Orders
   □ Order Template
   □ Automation Settings
   
2. TOOLS MENU:
   □ Order Macros
   □ API Settings
   □ External Order Interface
   □ Plugin Manager
   
3. TRADING MENU:
   □ Automated Trading
   □ Strategy Builder
   □ Order Hotkeys
   
4. RIGHT-CLICK ON CHART:
   □ Create Order
   □ Quick Buy/Sell
   □ Order from Excel
   
5. CHECK THE PDFs:
   - "R Trader Pro Trader's Guide.pdf"
   - "R Trader Trader's Guide.pdf"
   Look for chapters on:
   - Automation
   - Excel Integration
   - API/External Orders
   - Hotkeys
""")


if __name__ == "__main__":
    print("R|TRADER PRO ORDER EXECUTION INVESTIGATION\n")

    # Check Excel for order capabilities
    check_excel_macros()
    check_excel_formulas_for_orders()

    # Try COM interface
    test_rtrader_com_interface()

    # Show possibilities
    simulate_order_entry()

    # What to check manually
    check_rtrader_menus()

    print("\n" + "=" * 60)
    print("CONCLUSION")
    print("=" * 60)
    print("""
The Excel bridge is likely DATA-ONLY (no order execution).
For automated order execution, you'll need:

1. The official Rithmic API (wait for credentials)
2. Or find a hidden feature in R|Trader Pro menus
3. Or use manual execution based on Python signals

The safest approach: Python generates signals → You execute manually
""")

    input("\nPress Enter to exit...")
