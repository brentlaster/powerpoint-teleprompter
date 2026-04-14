#!/bin/bash
# Launch the teleprompter with ngrok tunnel for remote access.
# Use this when local networking won't work (hotel WiFi, etc.)
#
# Usage: ./launch3.sh [config.json]
#
# Requires: ngrok (brew install ngrok)
# You must have an ngrok account and have run 'ngrok config add-authtoken <token>' once.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# ── Check for ngrok ──
if ! command -v ngrok &>/dev/null; then
    echo "Error: ngrok is not installed."
    echo "Install with: brew install ngrok"
    echo "Then run: ngrok config add-authtoken <your-token>"
    echo "Get a free token at: https://dashboard.ngrok.com/get-started/your-authtoken"
    exit 1
fi

# ── Find config file ──
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

echo "Starting teleprompter with ngrok tunnel..."
echo "Config: $CONFIG"

# ── Read port from config ──
PORT=$(python3 -c "import json; print(json.load(open('$CONFIG')).get('port', 8765))" 2>/dev/null || echo 8765)

# ── Start ngrok in background ──
echo "Starting ngrok on port $PORT..."
ngrok http "$PORT" --log=stdout --log-level=warn > /tmp/ngrok-teleprompter.log 2>&1 &
NGROK_PID=$!

# Give ngrok a moment to start
sleep 2

# ── Get the public URL from ngrok's API ──
NGROK_URL=""
for i in $(seq 1 15); do
    NGROK_URL=$(curl -s http://localhost:4040/api/tunnels 2>/dev/null \
        | python3 -c "import sys,json; tunnels=json.load(sys.stdin).get('tunnels',[]); print(tunnels[0]['public_url'] if tunnels else '')" 2>/dev/null)
    if [ -n "$NGROK_URL" ]; then
        break
    fi
    sleep 1
done

if [ -z "$NGROK_URL" ]; then
    echo "Error: Could not get ngrok URL. Check if ngrok is authenticated."
    echo "Run: ngrok config add-authtoken <your-token>"
    echo "Log: /tmp/ngrok-teleprompter.log"
    kill $NGROK_PID 2>/dev/null
    exit 1
fi

echo ""
echo "══════════════════════════════════════════════════════"
echo "  ngrok tunnel active!"
echo "  Public URL:    $NGROK_URL"
echo "  Phone remote:  $NGROK_URL/remote"
echo "══════════════════════════════════════════════════════"
echo ""

# ── Cleanup function — kill ngrok when teleprompter exits ──
cleanup() {
    echo ""
    echo "Shutting down ngrok tunnel..."
    kill $NGROK_PID 2>/dev/null
    wait $NGROK_PID 2>/dev/null
    echo "Done."
}
trap cleanup EXIT INT TERM

# ── Wait for server to be ready, then open browser ──
(
  for i in $(seq 1 60); do
    if curl -s -o /dev/null "http://localhost:$PORT/api/state" 2>/dev/null; then
      open "http://localhost:$PORT" 2>/dev/null || xdg-open "http://localhost:$PORT" 2>/dev/null
      break
    fi
    sleep 0.5
  done
) &

# ── Launch teleprompter with the public URL ──
python3 "$SCRIPT_DIR/teleprompter.py" --config "$CONFIG" --no-browser --public-url "$NGROK_URL"
