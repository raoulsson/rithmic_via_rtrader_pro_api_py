import ctypes
import json
import mmap
import struct
import time


class RithmicSharedMemoryReader:
    def __init__(self):
        self.shared_memory_names = [
            'RithmicData',
            'RTraderData',
            'RithmicShared'
        ]
        self.mmaps = {}

    def read_shared_memory(self, name, size=65536):
        """Read from Windows shared memory"""
        print(f"\n{'=' * 60}")
        print(f"READING SHARED MEMORY: {name}")
        print(f"{'=' * 60}")

        try:
            # Open existing shared memory
            shm = mmap.mmap(-1, size, tagname=name, access=mmap.ACCESS_READ)

            # Read all data
            shm.seek(0)
            data = shm.read(size)

            # Find where data ends (look for null bytes)
            data_end = data.find(b'\x00' * 100)
            if data_end > 0:
                data = data[:data_end]

            print(f"Read {len(data)} bytes from {name}")

            # Analyze the data
            self.analyze_memory_structure(data)

            # Keep reference for monitoring
            self.mmaps[name] = shm

            return data

        except Exception as e:
            print(f"Error reading {name}: {e}")
            return None

    def analyze_memory_structure(self, data):
        """Analyze the structure of shared memory data"""
        if not data:
            return

        print("\nData Analysis:")
        print("-" * 40)

        # Check if it's text
        try:
            text = data.decode('utf-8', errors='ignore')
            if text.isprintable() or '\n' in text or '\r' in text:
                print("Text format detected:")
                print(text[:500])

                # Check for JSON
                if text.strip().startswith('{') or text.strip().startswith('['):
                    try:
                        parsed = json.loads(text.strip())
                        print("\nJSON structure found:")
                        print(json.dumps(parsed, indent=2)[:500])
                    except:
                        pass
                return
        except:
            pass

        # Check for binary structures
        print("Binary format detected")
        print(f"First 64 bytes (hex): {data[:64].hex()}")

        # Try to identify structure
        # Check for common patterns
        if data[:4] == b'RTHM':
            print("→ Rithmic header found: RTHM")

        # Try different interpretations
        if len(data) >= 8:
            # As integers
            int32_values = struct.unpack('ii', data[:8])
            print(f"As int32: {int32_values}")

            # As floats
            float_values = struct.unpack('ff', data[:8])
            print(f"As floats: {float_values}")

            # As doubles
            if len(data) >= 16:
                double_values = struct.unpack('dd', data[:16])
                print(f"As doubles: {double_values}")

        # Look for repeating structures
        self.find_repeating_patterns(data)

    def find_repeating_patterns(self, data):
        """Look for repeating structures in binary data"""
        print("\nSearching for patterns...")

        # Common structure sizes for market data
        sizes = [16, 20, 24, 32, 40, 48, 64, 80, 96, 128]

        for size in sizes:
            if len(data) >= size * 2:
                chunk1 = data[:size]
                chunk2 = data[size:size * 2]

                # Check if chunks have similar structure
                similarity = sum(1 for a, b in zip(chunk1, chunk2) if a == b)
                if similarity > size * 0.3:  # 30% similar
                    print(f"→ Possible {size}-byte structure detected")

    def monitor_changes(self, name, duration=10):
        """Monitor shared memory for changes"""
        print(f"\nMonitoring {name} for {duration} seconds...")
        print("-" * 40)

        if name not in self.mmaps:
            print(f"First read {name} before monitoring")
            return

        shm = self.mmaps[name]
        previous = shm.read()
        shm.seek(0)

        start_time = time.time()
        changes_detected = 0

        while time.time() - start_time < duration:
            time.sleep(0.1)

            shm.seek(0)
            current = shm.read()

            if current != previous:
                changes_detected += 1
                # Find what changed
                for i in range(min(len(current), len(previous))):
                    if current[i] != previous[i]:
                        print(f"Change at offset {i}: {previous[i:i + 10].hex()} → {current[i:i + 10].hex()}")
                        break

                previous = current

        print(f"Total changes detected: {changes_detected}")

    def dump_to_file(self, name):
        """Dump shared memory to file for analysis"""
        if name not in self.mmaps:
            print(f"Read {name} first")
            return

        filename = f"rithmic_{name}_{int(time.time())}.bin"
        shm = self.mmaps[name]
        shm.seek(0)
        data = shm.read()

        with open(filename, 'wb') as f:
            f.write(data)

        print(f"Dumped to {filename}")
        return filename

    def run_discovery(self):
        """Run shared memory discovery"""
        print("=" * 60)
        print("RITHMIC SHARED MEMORY DISCOVERY")
        print("=" * 60)

        # Read all shared memory segments
        for name in self.shared_memory_names:
            data = self.read_shared_memory(name)

            if data:
                # Monitor for changes
                self.monitor_changes(name, duration=5)

                # Dump to file
                self.dump_to_file(name)


class WiresharkGuide:
    @staticmethod
    def print_capture_guide():
        print("\n" + "=" * 60)
        print("WIRESHARK CAPTURE GUIDE FOR ATAS → RITHMIC")
        print("=" * 60)

        print("""
SETUP WIRESHARK:
----------------
1. Start Wireshark as Administrator
2. Select your network interface (Ethernet/WiFi)
3. Apply capture filter: "tcp port 8000 or tcp port 8100 or tcp port 8500"
   (Common Rithmic ports)

CAPTURE PROCESS:
----------------
1. Start capturing in Wireshark
2. Start ATAS and connect to Rithmic
3. Place a test order in ATAS
4. Stop capture after connection established

WIRESHARK FILTERS:
------------------
Display filters to find Rithmic traffic:
- tcp.port == 8000          # Common gateway port
- tcp.port == 8100          # Market data port
- tcp.port == 8500          # Order entry port
- ip.dst == rithmic.com     # All Rithmic traffic
- tcp.flags.syn == 1        # Connection initiation
- frame contains "RITHMIC"  # If protocol has text

FIND THE PROTOCOL:
------------------
1. Right-click first packet → Follow → TCP Stream
2. Look for:
   - JSON: {"user":"xxx","password":"xxx"}
   - FIX: 8=FIX.4.4|35=A|
   - Binary: Look for patterns
   
3. Check packet details:
   - Is it TLS/SSL encrypted? (port 443 or "TLS" in protocol)
   - Plain text visible?
   - Binary with structure?

EXPORT FOR ANALYSIS:
--------------------
1. File → Export Specified Packets
2. Save as .pcapng
3. Also export as:
   - File → Export Packet Dissections → As Plain Text
   - Include packet bytes

WHAT TO LOOK FOR:
-----------------
In the TCP stream:
1. Authentication sequence
2. Subscribe/unsubscribe messages  
3. Market data format
4. Order message format
5. Heartbeat/keepalive pattern

QUICK WIN:
----------
If you see plain JSON or FIX protocol, you can replicate it directly!
If encrypted (TLS), you'll need official API/SDK.
""")


def test_market_data_structure():
    """Test reading market data from shared memory"""
    print("\n" + "=" * 60)
    print("TESTING MARKET DATA EXTRACTION")
    print("=" * 60)

    # Common market data structure
    class MarketData(ctypes.Structure):
        _fields_ = [
            ("symbol", ctypes.c_char * 16),
            ("bid", ctypes.c_double),
            ("ask", ctypes.c_double),
            ("last", ctypes.c_double),
            ("volume", ctypes.c_int32),
            ("timestamp", ctypes.c_int64),
        ]

    try:
        # Try to read as structured data
        shm = mmap.mmap(-1, ctypes.sizeof(MarketData) * 100,
                        tagname="RithmicData", access=mmap.ACCESS_READ)

        # Read as market data array
        shm.seek(0)
        raw_data = shm.read(ctypes.sizeof(MarketData))

        # Parse as structure
        data = MarketData.from_buffer_copy(raw_data)

        print(f"Symbol: {data.symbol}")
        print(f"Bid: {data.bid}")
        print(f"Ask: {data.ask}")
        print(f"Last: {data.last}")
        print(f"Volume: {data.volume}")

        shm.close()

    except Exception as e:
        print(f"Could not parse as MarketData structure: {e}")


if __name__ == "__main__":
    print("Choose option:")
    print("1. Read Shared Memory")
    print("2. Wireshark Capture Guide")
    print("3. Test Market Data Structure")
    print("4. All of the above")

    choice = input("\nEnter choice (1-4): ").strip()

    if choice in ['1', '4']:
        reader = RithmicSharedMemoryReader()
        reader.run_discovery()

    if choice in ['2', '4']:
        WiresharkGuide.print_capture_guide()

    if choice in ['3', '4']:
        test_market_data_structure()

    print("\nPress Enter to exit...")
    input()
