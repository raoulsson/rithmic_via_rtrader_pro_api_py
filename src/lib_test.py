print("Testing imports...")
all_good = True

try:
    import websocket
    print("✓ websocket-client installed")
except ImportError as e:
    print(f"✗ websocket-client NOT installed: {e}")
    all_good = False

try:
    import pandas as pd
    print("✓ pandas installed")
except ImportError as e:
    print(f"✗ pandas NOT installed: {e}")
    all_good = False

try:
    import numpy as np
    print("✓ numpy installed")
except ImportError as e:
    print(f"✗ numpy NOT installed: {e}")
    all_good = False

try:
    import win32com.client
    print("✓ pywin32 installed")
except ImportError as e:
    print(f"✗ pywin32 NOT installed: {e}")
    all_good = False

if all_good:
    print("\n✅ All imports successful! Ready to trade.")
else:
    print("\n❌ Some packages missing. Please install them first.")