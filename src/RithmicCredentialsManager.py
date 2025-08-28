import json
import os
import socket
import time
import winreg


class RithmicCredentialsManager:
    def __init__(self):
        self.ports = [3010, 3011, 3012, 3013]

        # Common credential fields for Rithmic
        self.rithmic_credentials = {
            # These are typically required for direct Rithmic API
            '****': '',  # Your username
            '****': '',  # Your password
            'system_name': '',  # System identifier (assigned by Rithmic)
            'fcm_id': '',  # FCM (Futures Commission Merchant) ID
            'ib_id': '',  # IB (Introducing Broker) ID
            'environment': 'TEST',  # TEST or LIVE
            'gateway': '',  # Gateway server
        }

        # For local R|Trader Pro plugin connection
        self.plugin_credentials = {
            'plugin_name': 'PythonPlugin',
            'api_key': '',  # May need to generate in R|Trader
            'session_token': '',  # Might be available from current session
            'app_id': 'python_trader_1.0',
        }

    def find_rtrader_config(self):
        """Look for R|Trader Pro configuration files with credentials/settings"""
        print("Searching for R|Trader Pro configuration...")
        print("-" * 50)

        config_paths = [
            os.path.expanduser("~\\AppData\\Roaming\\Rithmic"),
            os.path.expanduser("~\\AppData\\Local\\Rithmic"),
            os.path.expanduser("~\\Documents\\Rithmic"),
            "C:\\ProgramData\\Rithmic",
            "C:\\Program Files (x86)\\Rithmic\\Rithmic Trader Pro",
        ]

        config_files = []
        for base_path in config_paths:
            if os.path.exists(base_path):
                print(f"✓ Found path: {base_path}")
                for root, dirs, files in os.walk(base_path):
                    for file in files:
                        if file.endswith(('.ini', '.config', '.cfg', '.xml', '.json', '.settings')):
                            full_path = os.path.join(root, file)
                            config_files.append(full_path)
                            print(f"  → Config file: {file}")

                            # Try to read non-sensitive info
                            self._examine_config(full_path)

        return config_files

    def _examine_config(self, filepath):
        """Safely examine config file for connection info (not passwords)"""
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(1000)  # Read first 1KB

                # Look for connection-related settings (not passwords)
                keywords = ['server', 'host', 'port', 'gateway', 'api', 'plugin', 'fcm', 'environment']
                for keyword in keywords:
                    if keyword.lower() in content.lower():
                        print(f"    Contains '{keyword}' configuration")

        except Exception as e:
            pass

    def check_registry_credentials(self):
        """Check Windows Registry for stored settings"""
        print("\nChecking Windows Registry for settings...")
        print("-" * 50)

        try:
            # Common registry locations for app settings
            paths = [
                (winreg.HKEY_CURRENT_USER, r"Software\Rithmic"),
                (winreg.HKEY_CURRENT_USER, r"Software\RithmicTrader"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Rithmic"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Rithmic"),
            ]

            for hkey, subkey in paths:
                try:
                    key = winreg.OpenKey(hkey, subkey)
                    print(f"✓ Found registry key: {subkey}")

                    # Enumerate values
                    i = 0
                    while True:
                        try:
                            name, value, type = winreg.EnumValue(key, i)
                            # Only show non-sensitive values
                            if 'password' not in name.lower() and 'pwd' not in name.lower():
                                print(f"  → {name}: {value}")
                            i += 1
                        except WindowsError:
                            break

                    winreg.CloseKey(key)
                except:
                    continue

        except Exception as e:
            print(f"Registry check error: {e}")

    def test_auth_methods(self, port):
        """Test different authentication methods"""
        print(f"\nTesting authentication methods on port {port}...")
        print("-" * 50)

        auth_messages = [
            # JSON-based auth
            {
                "type": "auth",
                "credentials": {
                    "app": "PythonTrader",
                    "version": "1.0"
                }
            },
            {
                "action": "login",
                "username": "plugin",
                "client_id": "python_client"
            },
            {
                "command": "authenticate",
                "api_key": "local_plugin",
                "timestamp": int(time.time())
            },
            # Session-based (if R|Trader shares its session)
            {
                "type": "attach_session",
                "request_session": True
            },
            {
                "msg": "get_session_token"
            },
            # Plugin registration
            {
                "register_plugin": {
                    "name": "PythonPlugin",
                    "type": "data_consumer",
                    "version": "1.0"
                }
            }
        ]

        for auth_msg in auth_messages:
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                sock.connect(('127.0.0.1', port))

                msg = json.dumps(auth_msg) + '\n'
                sock.send(msg.encode())

                response = sock.recv(4096)
                if response:
                    print(f"  Response to {auth_msg.get('type') or auth_msg.get('action')}:")
                    print(f"    {response.decode('utf-8', errors='ignore')[:100]}")

                sock.close()

            except Exception as e:
                continue

    def generate_credential_templates(self):
        """Generate credential configuration templates"""
        print("\n" + "=" * 60)
        print("CREDENTIAL TEMPLATES")
        print("=" * 60)

        print("\n1. For Direct Rithmic API Connection (like ATAS uses):")
        print("-" * 50)

        rithmic_config = """
# rithmic_credentials.py
RITHMIC_CONFIG = {
    'user': 'YOUR_USERNAME',
    'password': 'YOUR_PASSWORD',
    'system_name': 'SYSTEM_NAME',  # Provided by Rithmic
    'fcm_id': 'FCM_CODE',          # e.g., 'IronBeam', 'Advantage'
    'ib_id': 'IB_CODE',            # Your introducing broker
    'environment': 'TEST',          # or 'LIVE'
    
    # Server endpoints (examples)
    'gateway_host': 'gateway.rithmic.com',
    'gateway_port': 8000,
    
    # For market data
    'md_host': 'md.rithmic.com',
    'md_port': 8100,
}

# These credentials are what ATAS/QuantTower use
# Contact your broker for these values
"""
        print(rithmic_config)

        print("\n2. For Local R|Trader Pro Plugin Connection:")
        print("-" * 50)

        plugin_config = """
# local_plugin_config.py
PLUGIN_CONFIG = {
    # Option 1: No auth needed (local trusted connection)
    'connection': {
        'host': '127.0.0.1',
        'port': 3010,
        'trusted': True
    },
    
    # Option 2: API Key (generate in R|Trader Pro)
    'api_key': 'CHECK_RTRADER_SETTINGS_FOR_THIS',
    
    # Option 3: Windows Authentication (current user)
    'use_windows_auth': True,
    
    # Option 4: Session sharing (piggyback on R|Trader session)
    'use_existing_session': True,
}

# Check R|Trader Pro menus:
# - File → Preferences → API/Plugins
# - Tools → Generate API Key
# - View → Plugin Settings
"""
        print(plugin_config)

        print("\n3. Environment Variables Option:")
        print("-" * 50)

        env_config = """
# Set in Windows environment variables or .env file

# For Rithmic Direct
RITHMIC_USER=your_username
RITHMIC_PASSWORD=your_password
RITHMIC_FCM=your_fcm
RITHMIC_IB=your_ib
RITHMIC_SYSTEM=your_system_name

# For Local Plugin
RTRADER_API_KEY=your_api_key
RTRADER_PORT=3010
"""
        print(env_config)

    def check_rtrader_session(self):
        """Try to get session info from running R|Trader Pro"""
        print("\nChecking for R|Trader Pro session info...")
        print("-" * 50)

        # Check if we can get process info
        try:
            import psutil

            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                if 'rithmic' in proc.info['name'].lower():
                    print(f"Found process: {proc.info['name']} (PID: {proc.info['pid']})")

                    # Check command line for any useful info
                    if proc.info['cmdline']:
                        for arg in proc.info['cmdline']:
                            if '=' in arg and 'password' not in arg.lower():
                                print(f"  → Arg: {arg}")

        except ImportError:
            print("Install psutil for better process inspection: pip install psutil")
        except Exception as e:
            print(f"Process check error: {e}")

    def run_credential_discovery(self):
        """Run full credential discovery"""
        print("=" * 60)
        print("RITHMIC CREDENTIAL DISCOVERY")
        print("=" * 60)

        # Find config files
        config_files = self.find_rtrader_config()

        # Check registry
        self.check_registry_credentials()

        # Check running session
        self.check_rtrader_session()

        # Test auth on each port
        for port in self.ports:
            self.test_auth_methods(port)

        # Generate templates
        self.generate_credential_templates()

        print("\n" + "=" * 60)
        print("NEXT STEPS FOR CREDENTIALS")
        print("=" * 60)
        print("""
1. CHECK R|TRADER PRO:
   - Look for "API Key" or "Plugin Settings" in menus
   - Check Help → API Documentation
   - Look for "Generate Token" option

2. CONTACT YOUR BROKER:
   - They provide Rithmic API credentials
   - Different from your trading login
   - Includes: FCM ID, IB ID, System Name

3. TRY NO-AUTH LOCAL CONNECTION:
   - R|Trader might trust local connections
   - Try connecting without credentials first

4. USE WIRESHARK:
   - Capture what ATAS sends when connecting
   - Filter: tcp.port == 3010
   - Look for auth handshake

5. CHECK DOCUMENTATION:
   - %PROGRAMFILES%\\Rithmic\\Rithmic Trader Pro\\
   - Look for .pdf, .txt, .html files
   - Search for "API" or "Plugin" docs
""")


if __name__ == "__main__":
    manager = RithmicCredentialsManager()
    manager.run_credential_discovery()

    print("\nPress Enter to exit...")
    input()
