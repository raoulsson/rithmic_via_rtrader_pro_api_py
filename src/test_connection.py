import socket
import websocket
import json
import time
import win32com.client
import pythoncom
from typing import Optional, Dict, Any

class RTraderProTester:
    def __init__(self):
        self.rtd_progid = "RithmicRTD.RTD"  # Common ProgID for Rithmic RTD
        self.ws_url = "ws://127.0.0.1:8000/rithmic"
        self.port = 8000
        self.rtd = None

    def test_port_open(self) -> bool:
        """Check if R|Trader Pro port is accessible"""
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)
        result = sock.connect_ex(('127.0.0.1', self.port))
        sock.close()

        if result == 0:
            print(f"✓ Port {self.port} is open - R|Trader Pro appears to be running")
            return True
        else:
            print(f"✗ Port {self.port} not accessible")
            print("  → Start R|Trader Pro and enable 'Allow Plug-ins'")
            return False

    def test_websocket(self) -> bool:
        """Test WebSocket connectivity"""
        try:
            ws = websocket.create_connection(self.ws_url, timeout=2)
            print(f"✓ WebSocket connection established to {self.ws_url}")

            # Try sending a test message
            test_msg = {"type": "PING", "timestamp": time.time()}
            ws.send(json.dumps(test_msg))

            # Try to receive response (with timeout)
            ws.settimeout(1)
            try:
                response = ws.recv()
                print(f"  → Received response: {response[:100]}...")
            except:
                print("  → No immediate response (this may be normal)")

            ws.close()
            return True

        except ConnectionRefusedError:
            print(f"✗ WebSocket connection refused at {self.ws_url}")
            print("  → R|Trader Pro may not have plugins enabled")
            return False
        except Exception as e:
            print(f"✗ WebSocket error: {type(e).__name__}: {e}")
            return False

    def test_rtd_registration(self) -> bool:
        """Check if Rithmic RTD is registered in Windows"""
        try:
            import winreg

            # Check both 32-bit and 64-bit registry locations
            locations = [
                (winreg.HKEY_CLASSES_ROOT, f"{self.rtd_progid}\\CLSID"),
                (winreg.HKEY_CLASSES_ROOT, self.rtd_progid),
            ]

            found = False
            for hkey, subkey in locations:
                try:
                    with winreg.OpenKey(hkey, subkey) as key:
                        found = True
                        break
                except:
                    continue

            if found:
                print(f"✓ RTD Server '{self.rtd_progid}' is registered in Windows")
                return True
            else:
                print(f"✗ RTD Server '{self.rtd_progid}' not found in registry")
                print("  → You may need to register the RTD server or install Rithmic RTD")
                return False

        except Exception as e:
            print(f"✗ Could not check RTD registration: {e}")
            return False

    def test_rtd_connection(self) -> bool:
        """Test actual RTD connection"""
        try:
            pythoncom.CoInitialize()

            # Try to create RTD object
            self.rtd = win32com.client.Dispatch(self.rtd_progid)
            print(f"✓ Successfully created RTD object '{self.rtd_progid}'")

            # Try to connect (ServerStart method)
            try:
                result = self.rtd.ServerStart(None)
                if result == 1:  # 1 typically means success
                    print("✓ RTD ServerStart successful")

                    # Try to get some test data
                    self.test_rtd_data_retrieval()

                    # Clean shutdown
                    self.rtd.ServerTerminate()
                    return True
                else:
                    print(f"✗ RTD ServerStart returned: {result}")
                    return False

            except Exception as e:
                print(f"✗ RTD ServerStart failed: {e}")
                return False

        except Exception as e:
            print(f"✗ Could not create RTD object: {e}")
            print("  → Make sure Rithmic RTD is installed and registered")
            print("  → You may need to run 'regsvr32 RithmicRTD.dll' as admin")
            return False
        finally:
            pythoncom.CoUninitialize()

    def test_rtd_data_retrieval(self):
        """Try to retrieve data via RTD"""
        if not self.rtd:
            return

        try:
            # Common RTD test - try to get connection status
            # Format: =RTD("RithmicRTD.RTD", "", "CONNECTION", "STATUS")
            topics = ["CONNECTION", "STATUS"]
            topic_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, topics)

            result = self.rtd.ConnectData(1, topic_array, True)
            print(f"  → RTD ConnectData result: {result}")

            # Try RefreshData
            topic_count = win32com.client.VARIANT(pythoncom.VT_I4 | pythoncom.VT_BYREF, 0)
            data_array = self.rtd.RefreshData(topic_count)

            if data_array:
                print(f"  → RTD RefreshData returned {topic_count.value} topics")

        except Exception as e:
            print(f"  → RTD data retrieval test: {e}")

    def test_excel_rtd_formula(self):
        """Generate Excel RTD formulas for testing"""
        print("\n📊 Excel RTD Formula Examples:")
        print("-" * 50)
        print("Paste these into Excel to test RTD:")
        print()
        print("Connection Status:")
        print(f'  =RTD("{self.rtd_progid}", "", "CONNECTION", "STATUS")')
        print()
        print("Market Data (replace SYMBOL with actual symbol):")
        print(f'  =RTD("{self.rtd_progid}", "", "SYMBOL", "LAST")')
        print(f'  =RTD("{self.rtd_progid}", "", "SYMBOL", "BID")')
        print(f'  =RTD("{self.rtd_progid}", "", "SYMBOL", "ASK")')
        print()
        print("Account Info:")
        print(f'  =RTD("{self.rtd_progid}", "", "ACCOUNT", "BALANCE")')
        print(f'  =RTD("{self.rtd_progid}", "", "ACCOUNT", "MARGIN")')

    def run_all_tests(self):
        """Run comprehensive test suite"""
        print("=" * 60)
        print("R|TRADER PRO RTD CONNECTION TEST SUITE")
        print("=" * 60)
        print()

        results = {
            "Port Open": False,
            "WebSocket": False,
            "RTD Registered": False,
            "RTD Connection": False
        }

        # Test 1: Port
        print("1. Testing Port Connectivity...")
        print("-" * 40)
        results["Port Open"] = self.test_port_open()
        print()

        # Test 2: WebSocket
        if results["Port Open"]:
            print("2. Testing WebSocket Connection...")
            print("-" * 40)
            results["WebSocket"] = self.test_websocket()
            print()

        # Test 3: RTD Registration
        print("3. Checking RTD Registration...")
        print("-" * 40)
        results["RTD Registered"] = self.test_rtd_registration()
        print()

        # Test 4: RTD Connection
        if results["RTD Registered"]:
            print("4. Testing RTD Connection...")
            print("-" * 40)
            results["RTD Connection"] = self.test_rtd_connection()
            print()

        # Summary
        print("=" * 60)
        print("TEST SUMMARY")
        print("=" * 60)
        for test, passed in results.items():
            status = "✓ PASS" if passed else "✗ FAIL"
            print(f"{test:20} {status}")

        # Excel formulas
        if results["RTD Registered"]:
            self.test_excel_rtd_formula()

        # Next steps
        print("\n" + "=" * 60)
        print("NEXT STEPS")
        print("=" * 60)

        if not results["Port Open"]:
            print("1. Start R|Trader Pro")
            print("2. Enable 'Allow Plug-ins' in settings")

        elif not results["WebSocket"]:
            print("1. Verify 'Allow Plug-ins' is enabled in R|Trader Pro")
            print("2. Check firewall settings for port 8000")

        elif not results["RTD Registered"]:
            print("1. Install Rithmic RTD if not already installed")
            print("2. Register the RTD DLL (run as Administrator):")
            print('   regsvr32 "C:\\Path\\To\\RithmicRTD.dll"')

        elif not results["RTD Connection"]:
            print("1. Ensure R|Trader Pro is logged in")
            print("2. Check RTD permissions in R|Trader Pro")
            print("3. Try the Excel formulas manually")

        else:
            print("✓ All tests passed! RTD interface appears to be ready.")
            print("1. Test the Excel formulas in a spreadsheet")
            print("2. Run your main trading application")

def test_alternate_rtd_progids():
    """Test alternative RTD ProgIDs that Rithmic might use"""
    alternate_ids = [
        "RithmicRTD.RTD",
        "Rithmic.RTD",
        "RTrader.RTD",
        "RITHMIC.RTD.1",
        "RithmicRTDServer.RTD"
    ]

    print("\nChecking for alternate RTD ProgIDs...")
    print("-" * 40)

    for prog_id in alternate_ids:
        try:
            pythoncom.CoInitialize()
            rtd = win32com.client.Dispatch(prog_id)
            print(f"✓ Found: {prog_id}")
            pythoncom.CoUninitialize()
            return prog_id
        except:
            continue

    print("✗ No alternate RTD ProgIDs found")
    return None

if __name__ == "__main__":
    # Run main test suite
    tester = RTraderProTester()

    # Check for alternate ProgIDs first
    alt_id = test_alternate_rtd_progids()
    if alt_id:
        tester.rtd_progid = alt_id

    # Run all tests
    tester.run_all_tests()

    # Keep window open
    input("\nPress Enter to exit...")