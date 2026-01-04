#!/bin/bash

echo "================================================================"
echo "   Pattern-Based Academic Document Formatter"
echo "   Backend Server Startup Script (Linux/Mac)"
echo "================================================================"
echo ""

# Navigate to script directory
cd "$(dirname "$0")"
cd backend

echo "[1/4] Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed"
    echo "Please install Python 3.8+ using your package manager"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-venv python3-pip"
    echo "  macOS: brew install python3"
    exit 1
fi

echo "[2/4] Setting up virtual environment..."
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to create virtual environment"
        exit 1
    fi
fi

echo "[3/4] Activating virtual environment and installing dependencies..."
source venv/bin/activate
pip install -r requirements.txt --quiet
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies"
    exit 1
fi

echo "[4/4] Creating required directories..."
mkdir -p uploads outputs

echo ""
echo "================================================================"
echo "   Server is starting..."
echo "   API will be available at: http://localhost:5000"
echo "   Health check: http://localhost:5000/health"
echo "   Press Ctrl+C to stop the server"
echo "================================================================"
echo ""

python3 pattern_formatter_backend.py
