import socket
import subprocess

import psutil


class ATASConnectionFinder:
    """Find ATAS network connections to help with Wireshark filtering"""

    def find_atas_connections(self):
        """Find all network connections from ATAS"""
        print("=" * 60)
        print("FINDING ATAS NETWORK CONNECTIONS")
        print("=" * 60)

        atas_processes = []
        rithmic_connections = []

        # Find ATAS process
        for proc in psutil.process_iter(['pid', 'name']):
            name = proc.info['name'].lower()
            if 'atas' in name or 'advanced' in name:
                atas_processes.append(proc.info)
                print(f"Found ATAS Process: {proc.info['name']} (PID: {proc.info['pid']})")

        if not atas_processes:
            print("ATAS not found running. Please start ATAS first.")
            return None

        # Get network connections for ATAS
        for proc_info in atas_processes:
            try:
                proc = psutil.Process(proc_info['pid'])
                connections = proc.connections(kind='tcp')

                print(f"\nConnections for {proc_info['name']}:")
                print("-" * 40)

                for conn in connections:
                    if conn.status == 'ESTABLISHED':
                        local = f"{conn.laddr.ip}:{conn.laddr.port}"
                        remote = f"{conn.raddr.ip}:{conn.raddr.port}" if conn.raddr else "N/A"

                        print(f"  {local} → {remote} ({conn.status})")

                        # Try to resolve hostname
                        if conn.raddr:
                            try:
                                hostname = socket.gethostbyaddr(conn.raddr.ip)[0]
                                print(f"    Host: {hostname}")

                                if 'rithmic' in hostname.lower():
                                    rithmic_connections.append({
                                        'ip': conn.raddr.ip,
                                        'port': conn.raddr.port,
                                        'hostname': hostname
                                    })
                            except:
                                pass

            except Exception as e:
                print(f"Error getting connections: {e}")

        return rithmic_connections

    def generate_wireshark_filters(self, connections):
        """Generate Wireshark filters based on found connections"""
        print("\n" + "=" * 60)
        print("WIRESHARK CAPTURE FILTERS")
        print("=" * 60)

        if not connections:
            # Generic Rithmic filters
            print("\nNo active Rithmic connections found. Use these generic filters:")
            print("-" * 40)
            filters = [
                "tcp port 8000 or tcp port 8100 or tcp port 8500",
                "host gateway.rithmic.com",
                "host 208.88.250.0/24",  # Common Rithmic IP range
                "tcp portrange 8000-8500",
            ]
        else:
            # Specific filters based on found connections
            print("\nBased on active connections, use these filters:")
            print("-" * 40)

            filters = []
            for conn in connections:
                filters.append(f"host {conn['ip']}")
                filters.append(f"tcp port {conn['port']}")
                filters.append(f"host {conn['ip']} and tcp port {conn['port']}")

        print("\n📋 CAPTURE FILTERS (use one of these):")
        for f in filters:
            print(f"  {f}")

        print("\n📋 DISPLAY FILTERS (after capture):")
        print("  tcp.flags.syn == 1           # See connection start")
        print("  tcp.len > 0                  # Only packets with data")
        print("  !tcp.analysis.retransmission # Hide retransmissions")

        return filters

    def print_capture_instructions(self):
        """Detailed capture instructions"""
        print("\n" + "=" * 60)
        print("STEP-BY-STEP CAPTURE INSTRUCTIONS")
        print("=" * 60)

        print("""
1. PREPARE WIRESHARK:
   ----------------
   • Run Wireshark as Administrator
   • Select your main network interface (Ethernet/WiFi)
   • Apply NO filter initially (capture everything)
   
2. CAPTURE SEQUENCE:
   ----------------
   • Start Wireshark capture (🔴 red button)
   • In ATAS: Disconnect from Rithmic (if connected)
   • Clear Wireshark display (optional)
   • In ATAS: Connect to Rithmic
   • Wait for connection to establish
   • In ATAS: Subscribe to MNQ data
   • Stop Wireshark capture
   
3. FIND RITHMIC TRAFFIC:
   ---------------------
   In Wireshark, try these display filters:
   • tcp.port == 8000
   • tcp.port == 8100  
   • tcp.port == 8500
   • frame contains "rithmic"
   • tcp.flags.syn == 1 and tcp.flags.ack == 0
   
4. ANALYZE THE PROTOCOL:
   ---------------------
   • Find a packet with data (tcp.len > 0)
   • Right-click → Follow → TCP Stream
   • Look at the data format:
     - Red = Client → Server (ATAS → Rithmic)
     - Blue = Server → Client (Rithmic → ATAS)
   
5. IDENTIFY PROTOCOL TYPE:
   -----------------------
   ASCII/Text:  Readable text, JSON, or FIX
   Binary:      Random-looking characters
   Protobuf:    Binary with occasional text
   Encrypted:   Completely random (TLS/SSL)
   
6. EXPORT FOR ANALYSIS:
   --------------------
   • File → Export Specified Packets → All packets
   • Save as .pcapng
   • Share key packets: Authentication, Subscribe, Data
""")

    def check_common_ports(self):
        """Check which Rithmic ports are accessible"""
        print("\n" + "=" * 60)
        print("CHECKING COMMON RITHMIC PORTS")
        print("=" * 60)

        # Common Rithmic server IPs and ports
        servers = [
            ("gateway.rithmic.com", 8000, "Gateway"),
            ("gateway.rithmic.com", 8100, "Market Data"),
            ("gateway.rithmic.com", 8500, "Order Entry"),
            ("208.88.250.150", 8000, "Direct Gateway"),
            ("208.88.250.151", 8100, "Direct Market Data"),
        ]

        for host, port, desc in servers:
            try:
                # Try to resolve and connect
                ip = socket.gethostbyname(host)
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                result = sock.connect_ex((ip, port))
                sock.close()

                if result == 0:
                    print(f"✓ {desc}: {host}:{port} ({ip}) - OPEN")
                else:
                    print(f"✗ {desc}: {host}:{port} - CLOSED")

            except socket.gaierror:
                print(f"✗ {desc}: {host} - Cannot resolve")
            except Exception as e:
                print(f"✗ {desc}: {host}:{port} - Error: {e}")

    def netstat_analysis(self):
        """Use netstat to find connections"""
        print("\n" + "=" * 60)
        print("NETSTAT ANALYSIS")
        print("=" * 60)

        try:
            # Run netstat to find established connections
            result = subprocess.run(['netstat', '-n'],
                                    capture_output=True, text=True)

            lines = result.stdout.split('\n')

            rithmic_ips = ['208.88.250', '52.36', '54.187']  # Common Rithmic IP ranges

            print("\nPossible Rithmic connections:")
            print("-" * 40)

            for line in lines:
                if 'ESTABLISHED' in line:
                    for ip_prefix in rithmic_ips:
                        if ip_prefix in line:
                            print(f"  {line.strip()}")
                            break

                    # Also check for common ports
                    if any(f":{port}" in line for port in ['8000', '8100', '8500']):
                        print(f"  {line.strip()}")

        except Exception as e:
            print(f"Netstat error: {e}")

    def run_analysis(self):
        """Run complete analysis"""
        # Find ATAS connections
        connections = self.find_atas_connections()

        # Generate filters
        self.generate_wireshark_filters(connections)

        # Print instructions
        self.print_capture_instructions()

        # Check common ports
        self.check_common_ports()

        # Netstat analysis
        self.netstat_analysis()

        print("\n" + "=" * 60)
        print("QUICK WIRESHARK FILTER TO TRY")
        print("=" * 60)
        print("""
If ATAS is connected to Rithmic right now, use this Wireshark display filter
to find the traffic (after capturing without filter):

tcp and !tcp.port == 443 and tcp.len > 0 and (tcp.port >= 8000 and tcp.port <= 8500)

Or simply capture EVERYTHING for 10 seconds while connecting ATAS,
then look for packets to/from these IPs:
- 208.88.250.x (Rithmic servers)
- 52.36.x.x (AWS West)
- 54.187.x.x (AWS Oregon)
""")


if __name__ == "__main__":
    print("Make sure ATAS is running and connected to Rithmic")
    input("Press Enter to scan connections...")

    finder = ATASConnectionFinder()
    finder.run_analysis()

    print("\nDone! Check Wireshark with the suggested filters.")
    input("Press Enter to exit...")
