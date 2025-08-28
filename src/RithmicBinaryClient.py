import socket
import struct


class RithmicBinaryClient:
    def __init__(self):
        self.host = "38.65.210.71"
        self.port = 64100
        self.sock = None

    def connect(self):
        """Connect to Rithmic server"""
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((self.host, self.port))
        print(f"Connected to {self.host}:{self.port}")

    def send_message(self, msg_type, payload):
        """Send a message with length prefix"""
        # Build message: [4-byte length][2-byte type][payload]
        msg = struct.pack('>H', msg_type) + payload
        length = len(msg)

        # Send length prefix + message
        packet = struct.pack('>I', length) + msg
        self.sock.send(packet)
        print(f"Sent {len(packet)} bytes")

    def receive_message(self):
        """Receive a message"""
        # Read length prefix
        length_data = self.sock.recv(4)
        if len(length_data) < 4:
            return None

        length = struct.unpack('>I', length_data)[0]

        # Read message body
        data = self.sock.recv(length)
        return data

    def test_connection(self):
        """Test the connection"""
        try:
            self.connect()

            # Wait for any server messages
            self.sock.settimeout(2)
            try:
                data = self.receive_message()
                if data:
                    print(f"Received: {data.hex()}")
            except socket.timeout:
                print("No initial message from server")

            # Try sending a ping/heartbeat
            # Mimicking the captured message format
            self.send_message(0x4242, b'\x00\x00\x00\x02\x7f\xfe' +
                              b'\x00\x00\x00\x00\x7f\xff' +
                              b'\x00\x00\x00\x06mrv_lb')

            # Wait for response
            response = self.receive_message()
            if response:
                print(f"Response: {response.hex()}")

        except Exception as e:
            print(f"Error: {e}")
        finally:
            if self.sock:
                self.sock.close()


# Test
client = RithmicBinaryClient()
client.test_connection()
