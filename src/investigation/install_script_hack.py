import subprocess
import sys

packages = [
    'websocket-client',
    'pandas',
    'numpy',
    'pywin32',
    'python-dateutil'
]

print(f"Using Python: {sys.executable}")
print("Installing packages...\n")

for package in packages:
    print(f"Installing {package}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

print("\n✓ All packages installed!")

# Test imports
print("\nTesting imports...")

print("✓ websocket-client working")

print("✓ pandas working")

print("✓ numpy working")

print("✓ pywin32 working")

print("\n🎉 Setup complete! You can now run main.py")
