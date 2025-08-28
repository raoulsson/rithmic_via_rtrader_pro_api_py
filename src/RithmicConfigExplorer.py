import os
import xml.etree.ElementTree as ET


class RithmicConfigExplorer:
    def __init__(self):
        self.config_paths = {
            'exe_config': r'C:\Program Files (x86)\Rithmic\Rithmic Trader Pro\Rithmic Trader Pro.exe.config',
            'appdata': r'C:\Users\raoul\AppData\Roaming\Rithmic',
            'localappdata': r'C:\Users\raoul\AppData\Local\Rithmic',
        }

    def read_exe_config(self):
        """Read the .exe.config XML file"""
        print("=" * 60)
        print("READING RITHMIC TRADER PRO.EXE.CONFIG")
        print("=" * 60)

        config_file = self.config_paths['exe_config']
        if os.path.exists(config_file):
            try:
                tree = ET.parse(config_file)
                root = tree.getroot()

                print(f"\nFound config: {config_file}\n")

                # Look for appSettings
                app_settings = root.find('.//appSettings')
                if app_settings:
                    print("App Settings:")
                    print("-" * 40)
                    for setting in app_settings.findall('add'):
                        key = setting.get('key')
                        value = setting.get('value')
                        # Don't print passwords
                        if key and 'password' not in key.lower():
                            print(f"  {key}: {value}")

                # Look for connection strings
                conn_strings = root.find('.//connectionStrings')
                if conn_strings:
                    print("\nConnection Strings:")
                    print("-" * 40)
                    for conn in conn_strings.findall('add'):
                        name = conn.get('name')
                        conn_str = conn.get('connectionString')
                        if name:
                            print(f"  {name}: {conn_str[:50]}...")

                # Look for system.serviceModel (WCF endpoints)
                service_model = root.find('.//{urn:schemas-microsoft-com:asm.v1}system.serviceModel')
                if service_model:
                    print("\nService Endpoints:")
                    print("-" * 40)
                    for endpoint in service_model.findall('.//endpoint'):
                        address = endpoint.get('address')
                        binding = endpoint.get('binding')
                        if address:
                            print(f"  {binding}: {address}")

                # Look for any custom sections
                print("\nOther Configuration Sections:")
                print("-" * 40)
                for child in root:
                    if child.tag not in ['appSettings', 'connectionStrings']:
                        print(f"  Section: {child.tag}")

            except Exception as e:
                print(f"Error reading config: {e}")
        else:
            print(f"Config file not found: {config_file}")

    def scan_appdata_folder(self):
        """Scan AppData folder for useful files"""
        print("\n" + "=" * 60)
        print("SCANNING APPDATA FOLDERS")
        print("=" * 60)

        for name, path in [('Roaming', self.config_paths['appdata']),
                           ('Local', self.config_paths['localappdata'])]:
            if os.path.exists(path):
                print(f"\n{name}: {path}")
                print("-" * 40)

                for root, dirs, files in os.walk(path):
                    for file in files:
                        file_lower = file.lower()
                        full_path = os.path.join(root, file)

                        # Interesting file types
                        if file_lower.endswith(('.ini', '.cfg', '.config', '.xml', '.json', '.dat', '.settings')):
                            file_size = os.path.getsize(full_path)
                            print(f"  {file} ({file_size} bytes)")

                            # Try to peek at content
                            if file_size < 10000:  # Only small files
                                self._peek_file(full_path)

                        # Look for specific keywords in filenames
                        keywords = ['api', 'plugin', 'connection', 'session', 'token', 'auth']
                        if any(kw in file_lower for kw in keywords):
                            print(f"  → Interesting: {file}")

    def _peek_file(self, filepath):
        """Safely peek at file content"""
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(500)

                # Look for interesting patterns
                if 'port' in content.lower() or 'api' in content.lower() or 'plugin' in content.lower():
                    print(f"    → Contains API/Port configuration")

                    # Try to extract port numbers
                    import re
                    ports = re.findall(r'port["\s:=]+(\d{4,5})', content, re.IGNORECASE)
                    if ports:
                        print(f"    → Ports found: {ports}")

        except:
            pass

    def check_plugin_architecture(self):
        """Check if R|Trader uses a different plugin architecture"""
        print("\n" + "=" * 60)
        print("CHECKING PLUGIN ARCHITECTURE")
        print("=" * 60)

        print("\nPossible Connection Methods:")
        print("-" * 40)

        # Check for DLL plugins
        plugin_dir = r'C:\Program Files (x86)\Rithmic\Rithmic Trader Pro\Plugins'
        if os.path.exists(plugin_dir):
            print(f"✓ Plugin directory exists: {plugin_dir}")
            for file in os.listdir(plugin_dir):
                print(f"  → {file}")
        else:
            print("✗ No standard plugin directory")

        # Check for COM/OLE
        print("\nChecking COM/OLE Automation...")
        try:
            import win32com.client

            # Try common ProgIDs
            progids = [
                'Rithmic.Application',
                'RithmicTrader.Application',
                'RTrader.Application',
                'Rithmic.API',
            ]

            for progid in progids:
                try:
                    obj = win32com.client.Dispatch(progid)
                    print(f"✓ Found COM object: {progid}")
                    # List methods
                    for attr in dir(obj):
                        if not attr.startswith('_'):
                            print(f"    Method: {attr}")
                    return obj
                except:
                    continue

            print("✗ No COM automation objects found")

        except ImportError:
            print("pywin32 needed for COM testing")

        # Check for Named Pipes
        print("\nChecking for Named Pipes...")
        pipe_names = [
            r'\\.\pipe\Rithmic',
            r'\\.\pipe\RithmicTrader',
            r'\\.\pipe\RTrader',
            r'\\.\pipe\RithmicAPI',
        ]

        for pipe in pipe_names:
            try:
                handle = os.open(pipe, os.O_RDWR)
                os.close(handle)
                print(f"✓ Found pipe: {pipe}")
            except:
                continue

        # Check for shared memory
        print("\nChecking for Shared Memory/IPC...")
        try:
            import mmap

            # Common shared memory names
            shm_names = ['RithmicData', 'RTraderData', 'RithmicShared']
            for name in shm_names:
                try:
                    # Try to open existing shared memory
                    shm = mmap.mmap(-1, 1024, tagname=name)
                    print(f"✓ Found shared memory: {name}")
                    shm.close()
                except:
                    continue

        except:
            pass

    def suggest_alternative_approach(self):
        """Suggest alternative approaches based on findings"""
        print("\n" + "=" * 60)
        print("ALTERNATIVE APPROACH RECOMMENDATIONS")
        print("=" * 60)

        print("""
Based on the investigation, here are alternative approaches:

1. DIRECT RITHMIC API (Recommended):
   Since ATAS works, use the same approach:
   - Get Rithmic API credentials from your broker
   - Use Python Rithmic library or protocol
   - Connect directly to Rithmic servers (not R|Trader)
   
   Libraries to try:
   - pyrithmic (unofficial): pip install pyrithmic
   - Official Rithmic Python SDK (contact Rithmic)

2. EXCEL RTD BRIDGE:
   If RTD works in Excel, use Excel as a bridge:
   - Use xlwings or pywin32 to control Excel
   - Excel connects via RTD to get data
   - Python reads from Excel cells
   
   Example:
   ```python
   import xlwings as xw
   wb = xw.Book('RithmicData.xlsx')
   sheet = wb.sheets['Data']
   bid = sheet.range('A1').value  # =RTD("Rithmic.RTD","","MNQ","BID")
   ```

3. WINDOWS AUTOMATION:
   Control R|Trader Pro directly via UI automation:
   - Use pywinauto or pyautogui
   - Read data from screen
   - Send orders via UI
   
4. CHECK FOR HTTP/REST API:
   Some versions might have undocumented REST API:
   - Try http://127.0.0.1:3010/api
   - Check with browser dev tools while R|Trader runs

5. CONTACT RITHMIC SUPPORT:
   Ask specifically about:
   - R|Trader Pro Plugin SDK
   - Local API documentation  
   - Sample code for plugins
   - Python integration examples
""")

    def run_exploration(self):
        """Run complete exploration"""
        # Read main config
        self.read_exe_config()

        # Scan AppData
        self.scan_appdata_folder()

        # Check plugin architecture
        self.check_plugin_architecture()

        # Provide recommendations
        self.suggest_alternative_approach()


if __name__ == "__main__":
    explorer = RithmicConfigExplorer()
    explorer.run_exploration()

    print("\nPress Enter to exit...")
    input()
