#!/bin/bash
# Launch the teleprompter with a config file.
# Usage: ./launch.sh [config.json]
#
# If no config file is given, defaults to talk.json in the same directory.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG="${1:-$SCRIPT_DIR/talk.json}"

if [ ! -f "$CONFIG" ]; then
    echo "Config file not found: $CONFIG"
    echo "Usage: $0 [config.json]"
    exit 1
fi

echo "Starting teleprompter with config: $CONFIG"
python3 "$SCRIPT_DIR/teleprompter.py" --config "$CONFIG"
