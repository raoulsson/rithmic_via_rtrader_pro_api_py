import socket
import websocket
import json
import time
import os
import glob
import win32com.client
import pythoncom
from typing import Optional, Dict, Any

class RithmicConnectionTester:
    def __init__(self):
        # Updated ports based on scan results
        self.rithmic_ports = [3010, 3011, 3012, 3013]
        self.data_port = 5555
        self.rtd_progids = [
            "Rithmic.RTD",
            "RithmicRTD.RTD",
            "RithmicTrader.RTD",
            "RTrader.RTD",
            "Rithmic.ExcelRTD"
        ]
        self.installation_path = r"C:\Program Files (x86)\Rithmic\Rithmic Trader Pro"

    def test_rithmic_ports(self) -> Dict[int, bool]:
        """Test all Rithmic ports"""
        print("Testing Rithmic Trader Pro Ports...")
        print("-" * 40)

        results = {}
        for port in self.rithmic_ports + [self.data_port]:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex(('127.0.0.1', port))
            sock.close()

            if result == 0:
                print(f"✓ Port {port} is OPEN")
                results[port] = True

                # Try to identify the service
                self._probe_port_protocol(port)
            else:
                print(f"✗ Port {port} is CLOSED")
                results[port] = False

        return results

    def _probe_port_protocol(self, port):
        """Try different protocols on the port"""
        # Test 1: Try WebSocket
        try:
            ws = websocket.create_connection(f"ws://127.0.0.1:{port}/", timeout=1)
            print(f"  → Port {port}: WebSocket connection successful!")
            ws.close()
            return "WebSocket"
        except:
            pass

        # Test 2: Try WebSocket with different paths
        paths = ["/rithmic", "/api", "/rtd", "/data", "/trading", "/"]
        for path in paths:
            try:
                ws = websocket.create_connection(f"ws://127.0.0.1:{port}{path}", timeout=0.5)
                print(f"  → Port {port}: WebSocket works at path '{path}'")
                ws.close()
                return f"WebSocket{path}"
            except:
                continue

        # Test 3: Try raw socket with Rithmic protocol hint
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            sock.connect(('127.0.0.1', port))

            # Send a simple message
            test_messages = [
                b"PING\n",
                b'{"type":"ping"}\n',
                b'\x00\x00\x00\x04PING',  # Length-prefixed
            ]

            for msg in test_messages:
                try:
                    sock.send(msg)
                    sock.settimeout(0.5)
                    response = sock.recv(1024)
                    if response:
                        print(f"  → Port {port}: Got response with test message")
                        print(f"     Response (first 50 bytes): {response[:50]}")
                        sock.close()
                        return "Custom Protocol"
                except:
                    continue

            sock.close()
        except:
            pass

        print(f"  → Port {port}: Protocol unknown (may need authentication)")
        return "Unknown"

    def find_rtd_files(self):
        """Search for RTD-related files"""
        print("\nSearching for RTD Components...")
        print("-" * 40)

        rtd_files = []
        dll_files = []

        # Check Rithmic installation directory
        if os.path.exists(self.installation_path):
            print(f"✓ Found installation: {self.installation_path}")

            for root, dirs, files in os.walk(self.installation_path):
                for file in files:
                    file_lower = file.lower()
                    full_path = os.path.join(root, file)

                    if 'rtd' in file_lower:
                        rtd_files.append(full_path)
                        print(f"  → RTD File: {file}")
                    elif 'excel' in file_lower:
                        dll_files.append(full_path)
                        print(f"  → Excel Integration: {file}")
                    elif file_lower.endswith('.dll') and any(x in file_lower for x in ['api', 'plugin', 'external']):
                        dll_files.append(full_path)
                        print(f"  → API/Plugin DLL: {file}")

        # Check for registered RTD servers
        print("\nChecking Windows Registry for RTD...")
        self._check_registry_rtd()

        return rtd_files, dll_files

    def _check_registry_rtd(self):
        """Check registry for RTD registrations"""
        try:
            import winreg

            # Check HKCR for RTD servers
            print("Scanning for registered RTD servers...")

            for progid in self.rtd_progids:
                try:
                    key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid)
                    print(f"  ✓ Found registered: {progid}")
                    winreg.CloseKey(key)

                    # Try to get CLSID
                    try:
                        clsid_key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID")
                        clsid = winreg.QueryValueEx(clsid_key, "")[0]
                        print(f"    CLSID: {clsid}")
                        winreg.CloseKey(clsid_key)
                    except:
                        pass

                except:
                    continue

            # Look for any Rithmic entries
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "")
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        if 'rithmic' in subkey_name.lower() or 'rtrader' in subkey_name.lower():
                            print(f"  → Found registry entry: {subkey_name}")
                        i += 1
                    except WindowsError:
                        break
                winreg.CloseKey(key)
            except:
                pass

        except Exception as e:
            print(f"  Registry scan error: {e}")

    def test_rtd_com(self):
        """Try to initialize RTD COM objects"""
        print("\nTesting RTD COM Objects...")
        print("-" * 40)

        pythoncom.CoInitialize()

        for progid in self.rtd_progids:
            try:
                rtd = win32com.client.Dispatch(progid)
                print(f"✓ Successfully created COM object: {progid}")

                # Try to start the server
                try:
                    result = rtd.ServerStart(None)
                    print(f"  → ServerStart returned: {result}")
                    rtd.ServerTerminate()
                except Exception as e:
                    print(f"  → ServerStart error: {e}")

                return progid

            except Exception as e:
                continue

        print("✗ No RTD COM objects could be created")
        print("  RTD may not be installed or registered")

        pythoncom.CoUninitialize()
        return None

    def generate_connection_code(self):
        """Generate connection code based on findings"""
        print("\n" + "=" * 60)
        print("SUGGESTED CONNECTION CODE")
        print("=" * 60)

        code = '''
# Based on scan results, try these connection parameters:

import socket
import json

class RithmicConnection:
    def __init__(self):
        # Rithmic Trader Pro uses these ports
        self.ports = {
            'primary': 3010,
            'secondary': 3011,
            'data1': 3012,
            'data2': 3013,
            'feed': 5555
        }
    
    def connect(self, port_type='primary'):
        port = self.ports.get(port_type, 3010)
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect(('127.0.0.1', port))
        print(f"Connected to Rithmic on port {port}")
        return sock

# For WebSocket connections (if supported):
# ws = websocket.create_connection("ws://127.0.0.1:3010/")

# Note: The actual protocol may require authentication
# Check Rithmic API documentation for message format
'''
        print(code)

    def run_all_tests(self):
        """Run comprehensive test suite"""
        print("=" * 60)
        print("RITHMIC TRADER PRO CONNECTION TEST")
        print("=" * 60)
        print()

        # Test 1: Port connectivity
        port_results = self.test_rithmic_ports()

        # Test 2: Find RTD files
        rtd_files, dll_files = self.find_rtd_files()

        # Test 3: Test RTD COM
        working_progid = self.test_rtd_com()

        # Generate connection code
        self.generate_connection_code()

        # Excel setup instructions
        print("\n" + "=" * 60)
        print("EXCEL RTD SETUP")
        print("=" * 60)

        if not rtd_files and not working_progid:
            print("⚠ RTD components not found. You may need to:")
            print("1. Download Rithmic Excel RTD Add-in separately")
            print("2. Check Rithmic's website or support for RTD installer")
            print("3. In Excel, go to File → Options → Add-ins")
            print("4. Look for 'Rithmic RTD' or similar")
        else:
            print("Try these Excel formulas:")
            progid = working_progid or "Rithmic.RTD"
            print(f'=RTD("{progid}","","CONNECTION","STATUS")')
            print(f'=RTD("{progid}","","MNQ","LAST")')
            print(f'=RTD("{progid}","","MNQ","BID")')
            print(f'=RTD("{progid}","","MNQ","ASK")')

        print("\n" + "=" * 60)
        print("NEXT STEPS")
        print("=" * 60)
        print("1. The API appears to use ports 3010-3013, not 8000")
        print("2. Update your connection code to use these ports")
        print("3. Check Rithmic documentation for the protocol format")
        print("4. For RTD: May need separate RTD add-in installation")

if __name__ == "__main__":
    tester = RithmicConnectionTester()
    tester.run_all_tests()

    input("\nPress Enter to exit...")