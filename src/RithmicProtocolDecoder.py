import socket
import struct
import time
from typing import Optional


class RithmicProtocolDecoder:
    """Decode and understand Rithmic's binary protocol"""

    @staticmethod
    def decode_packet(hex_str: str):
        """Decode a hex packet and show its structure"""
        data = bytes.fromhex(hex_str.replace(' ', ''))
        offset = 0

        # Length prefix
        msg_len = struct.unpack('>I', data[0:4])[0]
        print(f"Message Length: {msg_len}")
        offset = 4

        # Message type
        msg_type = struct.unpack('>H', data[4:6])[0]
        print(f"Message Type: 0x{msg_type:04x} ({msg_type})")
        offset = 6

        # Parse fields
        fields = []
        while offset < len(data):
            # Field structure: 4 bytes header + data
            if offset + 4 > len(data):
                break

            field_header = struct.unpack('>I', data[offset:offset + 4])[0]
            field_type = field_header >> 16  # Upper 16 bits
            field_data_len = field_header & 0xFFFF  # Lower 16 bits
            offset += 4

            if field_type == 0x7FFE or field_type == 0x7FF0:  # Numeric field
                if field_data_len == 0:
                    # Length in next 4 bytes
                    actual_len = struct.unpack('>I', data[offset:offset + 4])[0]
                    offset += 4
                    field_data = data[offset:offset + actual_len]
                    offset += actual_len
                else:
                    field_data = data[offset:offset + field_data_len]
                    offset += field_data_len

                # Try to decode as string
                try:
                    value = field_data.decode('ascii')
                    fields.append((field_type, value))
                    print(f"  Field 0x{field_type:04x}: '{value}'")
                except:
                    # Numeric value
                    if len(field_data) == 4:
                        value = struct.unpack('>I', field_data)[0]
                    else:
                        value = field_data.hex()
                    fields.append((field_type, value))
                    print(f"  Field 0x{field_type:04x}: {value}")

            elif field_type == 0x7FFF:  # String field
                str_len = struct.unpack('>I', data[offset:offset + 4])[0]
                offset += 4
                string = data[offset:offset + str_len].decode('ascii')
                offset += str_len
                fields.append((field_type, string))
                print(f"  String 0x{field_type:04x}: '{string}'")

            elif field_type == 0x0000 and field_data_len == 0:
                # Another encoding format
                if offset + 4 <= len(data):
                    str_len = struct.unpack('>I', data[offset:offset + 4])[0]
                    offset += 4
                    if offset + str_len <= len(data):
                        string = data[offset:offset + str_len].decode('ascii')
                        offset += str_len
                        fields.append((0, string))
                        print(f"  String: '{string}'")
            else:
                # Unknown field type
                if field_data_len > 0 and offset + field_data_len <= len(data):
                    field_data = data[offset:offset + field_data_len]
                    offset += field_data_len
                    fields.append((field_type, field_data.hex()))
                    print(f"  Field 0x{field_type:04x}: {field_data.hex()}")

        return fields


class RithmicBinaryProtocol:
    """Implementation of Rithmic's binary protocol"""

    def __init__(self, host: str = "38.65.210.71", port: int = 64100):
        self.host = host
        self.port = port
        self.sock: Optional[socket.socket] = None
        self.message_type = 0x4242  # BB type from captures

    def connect(self) -> bool:
        """Connect to Rithmic server"""
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.settimeout(5)
            self.sock.connect((self.host, self.port))
            print(f"✓ Connected to {self.host}:{self.port}")
            return True
        except Exception as e:
            print(f"✗ Connection failed: {e}")
            return False

    def build_message(self, msg_type: int, fields: list) -> bytes:
        """Build a Rithmic protocol message"""
        # Start with message type
        message = struct.pack('>H', msg_type)

        # Add fields
        for field in fields:
            if isinstance(field, tuple):
                field_type, value = field

                if field_type == 0x7FFF:  # String field
                    # String encoding: type + 0000 + length + data
                    message += struct.pack('>I', (field_type << 16))
                    message += struct.pack('>I', len(value))
                    message += value.encode('ascii')

                elif field_type in [0x7FF0, 0x7FFE]:  # Numeric/data field
                    # Numeric encoding
                    message += struct.pack('>I', (field_type << 16))
                    if isinstance(value, str):
                        message += struct.pack('>I', len(value))
                        message += value.encode('ascii')
                    elif isinstance(value, int):
                        message += struct.pack('>I', 4)
                        message += struct.pack('>I', value)
                    else:
                        message += struct.pack('>I', len(value))
                        message += value

                elif field_type == 0x2710:  # Special field from login
                    message += struct.pack('>I', field_type)
                    message += struct.pack('>I', len(value))
                    message += value.encode('ascii')

                else:
                    # Generic field
                    message += struct.pack('>H', field_type)
                    if isinstance(value, str):
                        encoded = value.encode('ascii')
                        message += struct.pack('>H', len(encoded))
                        message += encoded
                    elif isinstance(value, bytes):
                        message += struct.pack('>H', len(value))
                        message += value
            else:
                # Simple string
                message += struct.pack('>I', len(field))
                message += field.encode('ascii')

        # Add length prefix
        length = len(message)
        return struct.pack('>I', length) + message

    def send_ping(self) -> bool:
        """Send a ping message (from Frame 792)"""
        # Build ping message: 0000001042420000000100000000000470696e67
        message = self.build_message(0x4242, [
            (0x0001, b'\x00\x00\x00\x00'),  # Field 1
            (0x0000, 'ping')  # Ping string
        ])

        # Or use exact bytes from capture
        exact_ping = bytes.fromhex("0000001042420000000100000000000470696e67")

        try:
            self.sock.send(exact_ping)
            print(f"→ Sent ping ({len(exact_ping)} bytes)")
            return True
        except Exception as e:
            print(f"✗ Send failed: {e}")
            return False

    def send_login_sequence(self) -> bool:
        """Send the login sequence (from Frame 833)"""
        # This is the login_agent_repository message
        hex_data = "0000004b4242000000042710000000176c6f67696e5f6167656e745f7265706f7369746f7279637ff00000000a313735363335373538377ff0000000063134333030300000000000066d72765f6c62"

        login_msg = bytes.fromhex(hex_data)

        try:
            self.sock.send(login_msg)
            print(f"→ Sent login_agent_repository ({len(login_msg)} bytes)")
            return True
        except Exception as e:
            print(f"✗ Send failed: {e}")
            return False

    def receive_message(self) -> Optional[bytes]:
        """Receive a message from server"""
        try:
            # Read length prefix
            length_data = self.sock.recv(4)
            if len(length_data) < 4:
                return None

            msg_length = struct.unpack('>I', length_data)[0]
            print(f"← Expecting {msg_length} bytes")

            # Read message body
            data = b''
            while len(data) < msg_length:
                chunk = self.sock.recv(msg_length - len(data))
                if not chunk:
                    break
                data += chunk

            print(f"← Received {len(data)} bytes")
            return data

        except socket.timeout:
            print("← Timeout waiting for response")
            return None
        except Exception as e:
            print(f"← Receive error: {e}")
            return None

    def test_connection_sequence(self):
        """Test the full connection sequence as captured"""
        print("\n" + "=" * 60)
        print("TESTING RITHMIC CONNECTION SEQUENCE")
        print("=" * 60)

        if not self.connect():
            return

        try:
            # Step 1: Send ping
            print("\n1. Sending ping...")
            if self.send_ping():
                response = self.receive_message()
                if response:
                    print(f"   Response: {response.hex()[:100]}...")
                    # Decode response
                    decoder = RithmicProtocolDecoder()
                    decoder.decode_packet(response.hex())

            time.sleep(0.5)

            # Step 2: Send login_agent_repository
            print("\n2. Sending login_agent_repository...")
            if self.send_login_sequence():
                response = self.receive_message()
                if response:
                    print(f"   Response length: {len(response)} bytes")
                    # The response should contain connection details
                    try:
                        text = response.decode('ascii', errors='ignore')
                        if '128.177.47.170' in text:
                            print("   ✓ Got connection details!")
                            print(f"   Contains: {text[:200]}")
                    except:
                        pass

            # Step 3: Keep connection alive
            print("\n3. Waiting for server messages...")
            self.sock.settimeout(2)
            for i in range(3):
                response = self.receive_message()
                if response:
                    print(f"   Message {i + 1}: {len(response)} bytes")

        except Exception as e:
            print(f"Error: {e}")

        finally:
            if self.sock:
                self.sock.close()
                print("\n✓ Connection closed")


def analyze_captures():
    """Analyze the captured packets"""
    print("=" * 60)
    print("ANALYZING CAPTURED PACKETS")
    print("=" * 60)

    decoder = RithmicProtocolDecoder()

    print("\nFrame 792 - Initial Ping:")
    print("-" * 40)
    decoder.decode_packet("0000001042420000000100000000000470696e67")

    print("\nFrame 831 - Server Response:")
    print("-" * 40)
    decoder.decode_packet(
        "0000002d4242000000037ffe0000000231347ffe0000000f756e6b6e6f776e20726571756573747fff0000000470696e67")

    print("\nFrame 833 - Login Agent Repository:")
    print("-" * 40)
    decoder.decode_packet(
        "0000004b4242000000042710000000176c6f67696e5f6167656e745f7265706f7369746f7279637ff00000000a313735363335373538377ff0000000063134333030300000000000066d72765f6c62")

    print("\n" + "=" * 60)
    print("KEY FINDINGS")
    print("=" * 60)
    print("""
1. Protocol Structure:
   - 4-byte length prefix (big-endian)
   - 2-byte message type (0x4242 = 'BB')
   - Fields with type codes (0x7FF0, 0x7FFE, 0x7FFF, etc.)
   
2. Connection Sequence:
   a) Send ping
   b) Get "unknown request" response
   c) Send login_agent_repository with timestamp
   d) Receive connection details (IPs, ports)
   
3. Important Fields:
   - login_agent_repository
   - Timestamp: 1756357587
   - Session ID: 143000
   - Server: mrv_lb (Market Recovery Load Balance)
   
4. The actual authentication happens AFTER this!
   You need to capture the real login with username/password.
""")


if __name__ == "__main__":
    print("Rithmic Protocol Analysis\n")

    # First analyze what we know
    analyze_captures()

    # Then test connection
    print("\n" + "=" * 60)
    input("Press Enter to test connection...")

    client = RithmicBinaryProtocol()
    client.test_connection_sequence()
