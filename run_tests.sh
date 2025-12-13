#!/bin/bash
# Test runner script for Robot Framework tests

set -e

# Activate virtual environment
if [ -d "venv" ]; then
    if [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
        source venv/Scripts/activate
    else
        source venv/bin/activate
    fi
else
    echo "Virtual environment not found. Please run setup.sh first."
    exit 1
fi

# Set default values
OUTPUT_DIR=${OUTPUT_DIR:-../reports}
LOG_LEVEL=${LOG_LEVEL:-INFO}

# Parse command line arguments
TEST_FILE=${1:-test_*.robot}
INCLUDE_TAGS=${2:-}
EXCLUDE_TAGS=${3:-}

# Change to tests directory (required for relative resource paths)
cd tests

# Build robot command
ROBOT_CMD="robot --outputdir ${OUTPUT_DIR} --loglevel ${LOG_LEVEL}"

# Add tags if specified
if [ -n "$INCLUDE_TAGS" ]; then
    ROBOT_CMD="${ROBOT_CMD} --include ${INCLUDE_TAGS}"
fi

if [ -n "$EXCLUDE_TAGS" ]; then
    ROBOT_CMD="${ROBOT_CMD} --exclude ${EXCLUDE_TAGS}"
fi

# Add test file
ROBOT_CMD="${ROBOT_CMD} ${TEST_FILE}"

echo "Running Robot Framework tests..."
echo "Command: ${ROBOT_CMD}"
echo ""

# Run tests
eval ${ROBOT_CMD}

# Check exit code
EXIT_CODE=$?

if [ $EXIT_CODE -eq 0 ]; then
    echo ""
    echo "Tests passed successfully!"
    echo "View report: ${OUTPUT_DIR}/report.html"
else
    echo ""
    echo "Tests failed with exit code: $EXIT_CODE"
    echo "View report: ${OUTPUT_DIR}/report.html"
    exit $EXIT_CODE
fi

