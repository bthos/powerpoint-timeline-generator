#!/bin/bash
# Setup script for Robot Framework testing environment

set -e

echo "Setting up Robot Framework testing environment..."

# Check Python version (try python3 first, then python)
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
elif command -v python &> /dev/null; then
    PYTHON_CMD=python
else
    echo "Error: Python not found. Please install Python 3.7+ and ensure it's in PATH."
    echo "Download from: https://www.python.org/downloads/"
    exit 1
fi

python_version=$($PYTHON_CMD --version 2>&1 | awk '{print $2}')
echo "Python version: $python_version"

# Create virtual environment
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    $PYTHON_CMD -m venv venv
else
    echo "Virtual environment already exists"
fi

# Activate virtual environment
echo "Activating virtual environment..."
if [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
    source venv/Scripts/activate
else
    source venv/bin/activate
fi

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install dependencies
echo "Installing dependencies..."
$PYTHON_CMD -m pip install --upgrade pip
$PYTHON_CMD -m pip install -r requirements.txt

# Create necessary directories
echo "Creating directories..."
mkdir -p reports
mkdir -p tests/data/baseline

# Copy .env.example to .env if it doesn't exist
if [ ! -f ".env" ]; then
    if [ -f ".env.example" ]; then
        echo "Creating .env file from .env.example..."
        cp .env.example .env
        echo "Please edit .env file with your configuration"
    fi
else
    echo ".env file already exists"
fi

echo ""
echo "Setup complete!"
echo ""
echo "To activate the virtual environment:"
if [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
    echo "  source venv/Scripts/activate"
else
    echo "  source venv/bin/activate"
fi
echo ""
echo "To run tests:"
echo "  ./run_tests.sh"
echo ""
echo "Or manually:"
echo "  cd tests && robot --outputdir ../reports test_*.robot"

