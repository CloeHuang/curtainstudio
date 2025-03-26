import subprocess
import sys

# List of required packages
required_packages = ["openpyxl", "xlsxwriter", "pandas"]

# Install packages if missing
for package in required_packages:
    subprocess.call([sys.executable, "-m", "pip", "install", package])

# Run main script
print("✅ All dependencies installed. Running OTB.py...\n")
subprocess.call([sys.executable, "OTB.py"])
