import socket
import struct
import json
import time
import threading
from datetime import datetime

class RTraderProtocolProbe:
    def __init__(self):
        self.ports = [3010, 3011, 3012, 3013, 5555]
        self.active_connections = []

    def test_json_rpc(self, port):
        """Test JSON-RPC protocol"""
        print(f"\n[Port {port}] Testing JSON-RPC...")

        test_messages = [
            # Standard JSON-RPC
            {"jsonrpc": "2.0", "method": "ping", "id": 1},
            {"jsonrpc": "2.0", "method": "getInfo", "id": 2},
            {"jsonrpc": "2.0", "method": "subscribe", "params": {"symbol": "MNQ"}, "id": 3},

            # Generic JSON
            {"type": "ping", "timestamp": time.time()},
            {"cmd": "subscribe", "symbol": "MNQ"},
            {"action": "getQuote", "symbol": "MNQ"},
            {"request": "marketData", "symbol": "MNQ"},

            # Rithmic-style
            {"msg": "login", "user": "plugin"},
            {"msg": "subscribe", "contract": "MNQ"},
            {"command": "GET_POSITIONS"},
        ]

        for msg in test_messages:
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                sock.connect(('127.0.0.1', port))

                # Try with newline terminator
                data = json.dumps(msg) + '\n'
                sock.send(data.encode())

                # Wait for response
                sock.settimeout(1)
                response = sock.recv(4096)

                if response:
                    print(f"  ✓ Got response for {msg.get('method') or msg.get('type') or msg.get('msg')}:")
                    try:
                        decoded = response.decode('utf-8', errors='ignore')
                        print(f"    Raw: {decoded[:100]}")
                        if decoded.strip():
                            parsed = json.loads(decoded.strip())
                            print(f"    Parsed: {parsed}")
                    except:
                        print(f"    Binary: {response[:50].hex()}")
                    sock.close()
                    return True

                sock.close()

            except socket.timeout:
                sock.close()
                continue
            except Exception as e:
                continue

        return False

    def test_binary_protocol(self, port):
        """Test binary/proprietary protocol"""
        print(f"\n[Port {port}] Testing Binary Protocols...")

        test_messages = [
            # FIX-style (Start of Header)
            b"8=FIX.4.4\x019=40\x0135=A\x0149=CLIENT\x0156=RTRADER\x0110=000\x01",

            # Length-prefixed
            struct.pack('>I', 4) + b"PING",
            struct.pack('<I', 4) + b"PING",
            struct.pack('>H', 4) + b"PING",

            # Simple commands
            b"PING\r\n",
            b"HELLO\r\n",
            b"CONNECT\r\n",
            b"LOGIN\r\n",

            # Binary markers
            b"\x00\x00\x00\x04PING",
            b"\x01\x00\x00\x00PING",
            b"\xFF\xFE" + b"PING",
            ]

        for msg in test_messages:
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                sock.connect(('127.0.0.1', port))

                sock.send(msg)
                sock.settimeout(1)
                response = sock.recv(4096)

                if response:
                    print(f"  ✓ Got response for binary message:")
                    print(f"    Hex: {response[:50].hex()}")
                    print(f"    ASCII: {response.decode('utf-8', errors='ignore')[:50]}")
                    sock.close()
                    return True

                sock.close()

            except:
                continue

        return False

    def listen_for_broadcasts(self, port, duration=5):
        """Listen passively for any data broadcasts"""
        print(f"\n[Port {port}] Listening for broadcasts ({duration}s)...")

        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(duration)
            sock.connect(('127.0.0.1', port))

            self.active_connections.append(sock)
            start = time.time()
            data_received = []

            while time.time() - start < duration:
                try:
                    sock.settimeout(0.5)
                    data = sock.recv(4096)
                    if data:
                        data_received.append(data)
                        print(f"  ✓ Received {len(data)} bytes")
                except socket.timeout:
                    continue
                except:
                    break

            sock.close()
            self.active_connections.remove(sock)

            if data_received:
                print(f"  Total data: {sum(len(d) for d in data_received)} bytes")
                # Analyze first chunk
                if data_received[0]:
                    self.analyze_data(data_received[0])
                return True

        except Exception as e:
            print(f"  Connection failed: {e}")

        return False

    def analyze_data(self, data):
        """Analyze received data to determine format"""
        print("\n  Data Analysis:")

        # Check if JSON
        try:
            decoded = data.decode('utf-8')
            if decoded[0] in '{[':
                parsed = json.loads(decoded)
                print(f"    → JSON detected: {parsed}")
                return
        except:
            pass

        # Check if text protocol
        try:
            decoded = data.decode('utf-8', errors='ignore')
            if all(c in string.printable for c in decoded):
                print(f"    → Text protocol: {decoded[:100]}")
                return
        except:
            pass

        # Binary analysis
        print(f"    → Binary protocol:")
        print(f"      First 4 bytes (hex): {data[:4].hex()}")
        print(f"      First 4 bytes (int): {struct.unpack('>I', data[:4])[0] if len(data)>=4 else 'N/A'}")

        # Look for patterns
        if data[:3] == b"FIX":
            print(f"      → Looks like FIX protocol")
        elif data[0:1] in [b'\x00', b'\x01', b'\xFF']:
            print(f"      → Binary protocol with header byte: {data[0]}")

    def test_http_endpoints(self, port):
        """Test HTTP/REST endpoints"""
        print(f"\n[Port {port}] Testing HTTP endpoints...")

        endpoints = [
            "GET / HTTP/1.1\r\nHost: localhost\r\n\r\n",
            "GET /api HTTP/1.1\r\nHost: localhost\r\n\r\n",
            "GET /api/positions HTTP/1.1\r\nHost: localhost\r\n\r\n",
            "GET /api/quotes/MNQ HTTP/1.1\r\nHost: localhost\r\n\r\n",
            "GET /plugin HTTP/1.1\r\nHost: localhost\r\n\r\n",
            "GET /rithmic HTTP/1.1\r\nHost: localhost\r\n\r\n",
        ]

        for endpoint in endpoints:
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                sock.connect(('127.0.0.1', port))

                sock.send(endpoint.encode())
                response = sock.recv(4096)

                if response and b"HTTP" in response:
                    print(f"  ✓ HTTP endpoint found!")
                    print(f"    Response: {response.decode('utf-8', errors='ignore')[:200]}")
                    sock.close()
                    return True

                sock.close()
            except:
                continue

        return False

    def keep_alive_test(self, port):
        """Test if connection stays alive and needs heartbeat"""
        print(f"\n[Port {port}] Testing connection persistence...")

        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.connect(('127.0.0.1', port))
            print(f"  ✓ Connected")

            # Wait and see if connection stays open
            time.sleep(2)

            # Try to send data
            try:
                sock.send(b"TEST\n")
                print(f"  ✓ Connection still alive after 2 seconds")

                # Check for response
                sock.settimeout(1)
                response = sock.recv(1024)
                if response:
                    print(f"  ✓ Got response: {response[:50]}")

            except:
                print(f"  ✗ Connection closed by server")

            sock.close()

        except Exception as e:
            print(f"  ✗ Connection failed: {e}")

    def run_discovery(self):
        """Run complete discovery on all ports"""
        print("=" * 60)
        print("R|TRADER PRO LOCAL API DISCOVERY")
        print("=" * 60)

        working_ports = []

        for port in self.ports:
            print(f"\n{'='*60}")
            print(f"TESTING PORT {port}")
            print(f"{'='*60}")

            # Quick connectivity check
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex(('127.0.0.1', port))
            sock.close()

            if result != 0:
                print(f"Port {port} is not open")
                continue

            print(f"Port {port} is OPEN - running protocol tests...")

            # Run tests
            results = {
                'json_rpc': self.test_json_rpc(port),
                'binary': self.test_binary_protocol(port),
                'http': self.test_http_endpoints(port),
                'broadcast': self.listen_for_broadcasts(port, 3),
            }

            # Keep-alive test
            self.keep_alive_test(port)

            if any(results.values()):
                working_ports.append((port, results))

        # Summary
        print("\n" + "=" * 60)
        print("DISCOVERY SUMMARY")
        print("=" * 60)

        if working_ports:
            print("\nResponsive ports found:")
            for port, results in working_ports:
                print(f"\nPort {port}:")
                for test, passed in results.items():
                    if passed:
                        print(f"  ✓ {test}")
        else:
            print("\nNo responsive protocols found. Possibilities:")
            print("1. Requires authentication/handshake first")
            print("2. Uses proprietary protocol needing documentation")
            print("3. Requires specific plugin registration")
            print("4. These ports are for internal use only")

        print("\n" + "=" * 60)
        print("RECOMMENDATIONS")
        print("=" * 60)
        print("1. Check R|Trader Pro for plugin/API documentation")
        print("2. Look for SDK or sample code from Rithmic")
        print("3. Contact Rithmic support for local API specs")
        print("4. Try Wireshark to capture traffic when other plugins connect")

import string  # Add this import at the top

if __name__ == "__main__":
    probe = RTraderProtocolProbe()
    probe.run_discovery()

    print("\nDone. Press Enter to exit...")
    input()