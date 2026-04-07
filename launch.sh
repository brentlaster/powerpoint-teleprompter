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

# Read port from config (default 8765)
PORT=$(python3 -c "import json; print(json.load(open('$CONFIG')).get('port', 8765))" 2>/dev/null || echo 8765)

# Wait for the server to be ready, then open the browser
(
  # Poll until the server responds (up to 30 seconds)
  for i in $(seq 1 60); do
    if curl -s -o /dev/null "http://localhost:$PORT/api/state" 2>/dev/null; then
      open "http://localhost:$PORT" 2>/dev/null || xdg-open "http://localhost:$PORT" 2>/dev/null
      break
    fi
    sleep 0.5
  done
) &

python3 "$SCRIPT_DIR/teleprompter.py" --config "$CONFIG" --no-browser
