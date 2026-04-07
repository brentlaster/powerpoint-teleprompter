#!/bin/bash
# Launch the teleprompter with a config file.
# Usage: ./launch.sh [config.json]
#
# If no config file is given, defaults to talk.json in the same directory.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# If no argument given, look for talk.json in the current directory first,
# then fall back to the script's own directory.
if [ -n "$1" ]; then
    CONFIG="$1"
elif [ -f "./talk.json" ]; then
    CONFIG="$(pwd)/talk.json"
else
    CONFIG="$SCRIPT_DIR/talk.json"
fi

if [ ! -f "$CONFIG" ]; then
    echo "Config file not found: $CONFIG"
    echo "Usage: $0 [config.json]"
    exit 1
fi

echo "Starting teleprompter with config: $CONFIG"
python3 "$SCRIPT_DIR/teleprompter.py" --config "$CONFIG"
