import socket
import subprocess
import re

def find_rtrader_process():
    """Find R|Trader Pro process and its ports using netstat"""
    print("Searching for R|Trader Pro processes and ports...")
    print("-" * 50)

    try:
        # Get all TCP connections
        result = subprocess.run(['netstat', '-ano'], capture_output=True, text=True)
        lines = result.stdout.split('\n')

        # Get process list
        proc_result = subprocess.run(['tasklist', '/FO', 'CSV'], capture_output=True, text=True)
        processes = {}
        for line in proc_result.stdout.split('\n')[1:]:  # Skip header
            if line:
                parts = line.split('","')
                if len(parts) >= 2:
                    name = parts[0].strip('"')
                    pid = parts[1].strip('"')
                    processes[pid] = name

        # Look for R|Trader related processes
        rtrader_pids = []
        for pid, name in processes.items():
            if 'trader' in name.lower() or 'rithmic' in name.lower() or 'r|trader' in name.lower():
                print(f"Found process: {name} (PID: {pid})")
                rtrader_pids.append(pid)

        if not rtrader_pids:
            print("No R|Trader Pro process found. Is it running?")
            # Still check for common ports
            print("\nChecking common trading platform ports anyway...")

        # Parse netstat output for listening ports
        listening_ports = []
        for line in lines:
            if 'LISTENING' in line and '127.0.0.1' in line or '0.0.0.0' in line:
                parts = line.split()
                if len(parts) >= 5:
                    address = parts[1]
                    pid = parts[-1]

                    # Extract port from address (format: 127.0.0.1:port)
                    if ':' in address:
                        port = address.split(':')[-1]

                        # Check if it's an R|Trader process or common ports
                        if pid in rtrader_pids:
                            proc_name = processes.get(pid, "Unknown")
                            print(f"  → R|Trader Port {port} (Process: {proc_name})")
                            listening_ports.append(int(port))
                        elif int(port) in [8000, 8001, 8080, 8081, 9000, 9001, 5000, 5001, 3000, 3001, 4000, 4001]:
                            proc_name = processes.get(pid, "Unknown")
                            print(f"  → Port {port} is open (Process: {proc_name}, PID: {pid})")
                            listening_ports.append(int(port))

        return listening_ports

    except Exception as e:
        print(f"Error scanning processes: {e}")
        return []

def scan_common_ports():
    """Scan common ports that trading platforms use"""
    common_ports = [
        (8000, "Common API port"),
        (8001, "Alternative API port"),
        (8080, "HTTP alternate"),
        (8081, "HTTP alternate 2"),
        (9000, "Common service port"),
        (9001, "Alternative service"),
        (5000, "Common dev port"),
        (5001, "Alternative dev port"),
        (3000, "Common web port"),
        (4000, "Common data port"),
        (4001, "Alternative data port"),
        (7496, "IB TWS API"),
        (7497, "IB TWS Paper"),
        (5555, "Data feed port"),
        (6666, "Alternative feed"),
        (12345, "Debug port"),
        (15555, "Rithmic specific"),
        (16555, "Rithmic alternative")
    ]

    print("\nScanning common trading platform ports...")
    print("-" * 50)

    open_ports = []
    for port, description in common_ports:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(0.1)  # Quick timeout
        result = sock.connect_ex(('127.0.0.1', port))
        sock.close()

        if result == 0:
            print(f"✓ Port {port:5} OPEN - {description}")
            open_ports.append(port)

    if not open_ports:
        print("✗ No common trading platform ports found open")

    return open_ports

def test_port_service(port):
    """Try to identify what service is running on a port"""
    print(f"\nTesting port {port} for service identification...")
    print("-" * 40)

    # Try HTTP
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)
        sock.connect(('127.0.0.1', port))
        sock.send(b"GET / HTTP/1.0\r\n\r\n")
        response = sock.recv(1024)
        sock.close()

        if response:
            print(f"HTTP Response preview: {response[:100]}")
            if b"HTTP" in response:
                print("  → Appears to be an HTTP service")
            return "HTTP"
    except:
        pass

    # Try WebSocket
    try:
        import websocket
        ws = websocket.create_connection(f"ws://127.0.0.1:{port}/", timeout=1)
        print(f"  → WebSocket connection successful on port {port}")
        ws.close()
        return "WebSocket"
    except:
        pass

    # Try raw connection
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)
        sock.connect(('127.0.0.1', port))
        sock.send(b"\n")
        response = sock.recv(1024)
        sock.close()

        if response:
            print(f"Raw Response: {response[:50]}")
            return "Unknown Service"
    except:
        pass

    print("  → Could not identify service type")
    return None

def check_rtd_files():
    """Check for RTD-related files in common locations"""
    print("\nSearching for RTD files...")
    print("-" * 50)

    import os
    import glob

    common_paths = [
        r"C:\Program Files\R*Trader*\**\*.dll",
        r"C:\Program Files (x86)\R*Trader*\**\*.dll",
        r"C:\Program Files\Rithmic\**\*.dll",
        r"C:\Program Files (x86)\Rithmic\**\*.dll",
        r"C:\RithmicTrader\**\*.dll",
        r"C:\R*Trader*\**\*.dll",
    ]

    rtd_files = []
    for pattern in common_paths:
        try:
            files = glob.glob(pattern, recursive=True)
            for file in files:
                if 'rtd' in file.lower() or 'excel' in file.lower():
                    rtd_files.append(file)
                    print(f"Found: {file}")
        except:
            continue

    if not rtd_files:
        print("✗ No RTD-related DLL files found")
        print("  The RTD component may need to be installed separately")

    return rtd_files

def main():
    print("=" * 60)
    print("R|TRADER PRO PORT AND SERVICE SCANNER")
    print("=" * 60)
    print()

    # Step 1: Find R|Trader process and its ports
    process_ports = find_rtrader_process()

    print()

    # Step 2: Scan common ports
    open_ports = scan_common_ports()

    # Step 3: Test identified ports
    all_ports = list(set(process_ports + open_ports))
    if all_ports:
        print(f"\nTesting {len(all_ports)} open port(s) for services...")
        for port in all_ports:
            test_port_service(port)

    # Step 4: Look for RTD files
    check_rtd_files()

    # Summary and recommendations
    print("\n" + "=" * 60)
    print("RECOMMENDATIONS")
    print("=" * 60)

    if not all_ports:
        print("1. Verify R|Trader Pro is running")
        print("2. Check Windows Firewall - it may be blocking connections")
        print("3. Run R|Trader Pro as Administrator")
        print("4. Look for a 'Plugins', 'API', or 'External Connections' setting")
    else:
        print(f"Found open ports: {all_ports}")
        print("1. Update the connection test to use one of these ports")
        print("2. The API may be on a different port than expected (8000)")

    print("\nFor RTD:")
    print("1. Check if there's a separate RTD installer from Rithmic")
    print("2. Look in R|Trader Pro installation folder for RTD-related files")
    print("3. Contact Rithmic support for RTD setup documentation")

if __name__ == "__main__":
    main()
    input("\nPress Enter to exit...")