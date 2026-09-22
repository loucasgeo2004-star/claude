#!/bin/bash

# Setup script for Excel Formula Re-Pointer
# This script installs the tool and makes it available system-wide

set -e

echo "=================================================="
echo "Excel Formula Re-Pointer - Setup Installer"
echo "=================================================="
echo ""

# Color codes
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Step 1: Check Python
echo -e "${BLUE}[1/5]${NC} Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Python 3 not found!${NC}"
    echo "Please install Python 3 from https://www.python.org"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo -e "${GREEN}✓ Python $PYTHON_VERSION found${NC}"
echo ""

# Step 2: Install openpyxl
echo -e "${BLUE}[2/5]${NC} Installing required library (openpyxl)..."
python3 -m pip install -q openpyxl 2>/dev/null || python3 -m pip install openpyxl
echo -e "${GREEN}✓ openpyxl installed${NC}"
echo ""

# Step 3: Create installation directory
echo -e "${BLUE}[3/5]${NC} Setting up installation directory..."
INSTALL_DIR="$HOME/.local/bin"
mkdir -p "$INSTALL_DIR"
echo -e "${GREEN}✓ Directory ready: $INSTALL_DIR${NC}"
echo ""

# Step 4: Copy script
echo -e "${BLUE}[4/5]${NC} Copying script..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cp "$SCRIPT_DIR/excel_formula_repointer.py" "$INSTALL_DIR/excel_formula_repointer"
chmod +x "$INSTALL_DIR/excel_formula_repointer"
echo -e "${GREEN}✓ Script installed to: $INSTALL_DIR/excel_formula_repointer${NC}"
echo ""

# Step 5: Add to PATH
echo -e "${BLUE}[5/5]${NC} Adding to PATH..."

SHELL_RC=""
if [[ "$SHELL" == *"zsh"* ]]; then
    SHELL_RC="$HOME/.zshrc"
elif [[ "$SHELL" == *"bash"* ]]; then
    SHELL_RC="$HOME/.bashrc"
fi

if [ -n "$SHELL_RC" ]; then
    # Check if already in PATH
    if grep -q "$INSTALL_DIR" "$SHELL_RC"; then
        echo -e "${GREEN}✓ Already in PATH${NC}"
    else
        echo "" >> "$SHELL_RC"
        echo "# Excel Formula Re-Pointer" >> "$SHELL_RC"
        echo "export PATH=\"\$PATH:$INSTALL_DIR\"" >> "$SHELL_RC"
        echo -e "${GREEN}✓ Added to PATH in $SHELL_RC${NC}"
        echo -e "${BLUE}Note: Run 'source $SHELL_RC' or restart terminal to apply${NC}"
    fi
else
    echo -e "${RED}⚠ Could not detect shell (bash/zsh)${NC}"
    echo "Manually add this to your shell profile:"
    echo "export PATH=\"\$PATH:$INSTALL_DIR\""
fi

echo ""
echo "=================================================="
echo -e "${GREEN}✅ Installation Complete!${NC}"
echo "=================================================="
echo ""
echo "You can now run the script from anywhere:"
echo -e "${BLUE}  excel_formula_repointer your_file.xlsx${NC}"
echo ""
echo "Examples:"
echo -e "${BLUE}  excel_formula_repointer ~/Documents/Trident_2026.xlsx${NC}"
echo -e "${BLUE}  excel_formula_repointer Accounting.xlsx${NC}"
echo ""
echo "For help:"
echo -e "${BLUE}  excel_formula_repointer --help${NC}"
echo ""
