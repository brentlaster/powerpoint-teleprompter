#!/usr/bin/env python3
"""
Slide Teleprompter — VBA event-driven.

A small VBA macro inside your PowerPoint fires on every slide change
and sends an HTTP request to this script's local server.  No file I/O,
no sandbox issues.

Setup (one-time per presentation):
    1.  Open your .pptx in PowerPoint
    2.  Tools → Macro → Visual Basic Editor  (or Alt+F11 / Opt+F11)
    3.  Insert → Module
    4.  Paste the VBA code printed at startup (or from vba_macro.txt)
    5.  Close the VB editor
    6.  Save as .pptm (macro-enabled) if you want to keep the macro

Usage:
    python3 teleprompter.py script.md

    Then start your slideshow in PowerPoint.  The teleprompter
    auto-updates on every slide change.

Fallback:
    python3 teleprompter.py script.md --keyboard
    Uses pynput for global keyboard monitoring (period/comma keys)
    instead of VBA.  Requires: pip3 install pynput

Your script file should use SLIDE markers to delimit sections, e.g.:
    ## SLIDE 1
    Welcome everyone...

    ## SLIDE 2
    Let me start by...
"""

import argparse
import http.server
import json
import os
import re
import sys
import threading
import time
import webbrowser
from pathlib import Path

# ── Config ──
DEFAULT_PORT = 8765

VBA_MACRO_TEMPLATE = r'''
' ── Teleprompter bridge ──
' Paste this into a new Module in your PowerPoint VBA editor.
' (Tools → Macro → Visual Basic Editor → Insert → Module)
'
' On every slide change it pings the local teleprompter server.
' No file I/O — works within PowerPoint's sandbox.

Private Sub NotifyTeleprompter(idx As Long, total As Long)
    Dim url As String
    url = "http://localhost:{{PORT}}/api/slide/" & idx & "/" & total

    ' Use MacScript to run curl in the background (non-blocking)
    On Error Resume Next
    MacScript "do shell script ""curl -s -m 1 " & url & " > /dev/null 2>&1 &"""
    On Error GoTo 0
End Sub

Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
    Dim idx As Long
    idx = SSW.View.Slide.SlideIndex

    Dim total As Long
    total = SSW.Presentation.Slides.Count

    NotifyTeleprompter idx, total
End Sub

Sub OnSlideShowTerminate(ByVal SSW As SlideShowWindow)
    On Error Resume Next
    MacScript "do shell script ""curl -s -m 1 http://localhost:{{PORT}}/api/stopped > /dev/null 2>&1 &"""
    On Error GoTo 0
End Sub
'''.strip()

# ── Global state ──
current_slide = 1
total_slides = 1
script_sections = []
slideshow_active = False
mode = "vba"  # "vba" or "keyboard"


def parse_script(filepath):
    """Parse a markdown script file into sections split by SLIDE markers."""
    text = Path(filepath).read_text(encoding="utf-8")

    # Match lines like:
    #   ## [SLIDE 1 — Title]
    #   ## SLIDE 2
    #   **SLIDE 3**
    #   SLIDE 11b — Some Title
    # Captures everything up to and including the SLIDE marker line.
    marker = re.compile(
        r"^[#*\-\s]*\[?\s*SLIDE[\s:]*\d+[a-zA-Z]?\b.*$",
        re.IGNORECASE | re.MULTILINE,
    )

    # Split on marker lines
    parts = marker.split(text)

    # parts[0] = text before first SLIDE marker (preamble)
    # parts[1..] = text after each SLIDE marker
    if len(parts) < 2:
        return [text.strip()]

    sections = []
    for part in parts[1:]:
        sections.append(part.strip())

    return sections


# ── Slide navigation ──

def _set_slide(n):
    global current_slide
    current_slide = max(1, min(n, total_slides))


def next_slide():
    _set_slide(current_slide + 1)


def prev_slide():
    _set_slide(current_slide - 1)


def _receive_slide(idx, tot):
    """Called when VBA macro reports a slide change."""
    global current_slide, total_slides, slideshow_active
    current_slide = idx
    if tot > 0:
        total_slides = tot
    slideshow_active = True


def _receive_stopped():
    """Called when VBA macro reports slideshow ended."""
    global slideshow_active
    slideshow_active = False


# ── Keyboard listener (fallback) ──

def start_keyboard_listener():
    """Listen for global key events to advance/retreat slides."""
    try:
        from pynput import keyboard
    except ImportError:
        print("\n" + "=" * 60)
        print("ERROR: pynput is required for --keyboard mode.")
        print("Install it with:  pip3 install pynput")
        print("=" * 60 + "\n")
        sys.exit(1)

    ADVANCE_KEY = '.'
    RETREAT_KEY = ','

    def on_press(key):
        try:
            ch = key.char
            if ch == ADVANCE_KEY:
                next_slide()
            elif ch == RETREAT_KEY:
                prev_slide()
        except AttributeError:
            if key == keyboard.Key.home:
                _set_slide(1)
            elif key == keyboard.Key.end:
                _set_slide(total_slides)

    listener = keyboard.Listener(on_press=on_press)
    listener.daemon = True
    listener.start()
    print("Global keyboard listener started")
    print(f"  [ {ADVANCE_KEY} ]  (period) = next script section")
    print(f"  [ {RETREAT_KEY} ]  (comma)  = previous script section")
    return listener


# ── HTTP Server ──

class TeleprompterHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        if self.path == "/":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))

        elif self.path == "/api/state":
            slide_idx = current_slide - 1
            section_idx = max(0, min(slide_idx, len(script_sections) - 1))
            script_text = script_sections[section_idx] if script_sections else ""

            state = {
                "slide": current_slide,
                "totalSlides": total_slides,
                "scriptHtml": md_to_html(script_text),
                "totalSections": len(script_sections),
                "active": slideshow_active,
                "mode": mode,
            }
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(json.dumps(state).encode("utf-8"))

        elif self.path.startswith("/api/slide/"):
            # Called by VBA macro: /api/slide/{idx}/{total}
            try:
                parts = self.path.split("/")
                idx = int(parts[3])
                tot = int(parts[4]) if len(parts) > 4 else 0
                _receive_slide(idx, tot)
                self.send_response(200)
                self.send_header("Content-Type", "text/plain")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b"ok")
                print(f"  Slide {idx}/{tot}")
            except Exception:
                self.send_response(400)
                self.end_headers()

        elif self.path == "/api/stopped":
            # Called by VBA when slideshow ends
            _receive_stopped()
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"ok")
            print("  Slideshow ended")

        elif self.path.startswith("/api/goto/"):
            try:
                slide_num = int(self.path.split("/")[-1])
                _set_slide(slide_num)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": True, "slide": current_slide}).encode())
            except Exception:
                self.send_response(400)
                self.end_headers()

        else:
            self.send_response(404)
            self.end_headers()


def md_to_html(text):
    """Minimal markdown to HTML conversion."""
    import html as html_mod
    text = html_mod.escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"\*(.+?)\*", r"<em>\1</em>", text)
    text = re.sub(r"_(.+?)_", r"<em>\1</em>", text)
    text = re.sub(
        r"`(.+?)`",
        r'<code style="background:rgba(255,255,255,0.1);padding:2px 6px;border-radius:3px;">\1</code>',
        text,
    )
    paragraphs = re.split(r"\n\s*\n", text)
    return "".join(f"<p>{p.replace(chr(10), '<br>')}</p>" for p in paragraphs)


# ── HTML Page ──

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Teleprompter</title>
<style>
  :root {
    --bg: #111118;
    --bg-panel: #1a1a28;
    --accent: #e94560;
    --text: #eaeaea;
    --dim: #777;
    --border: #2a2a3e;
    --green: #2ecc71;
    --orange: #e67e22;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
    background: var(--bg);
    color: var(--text);
    height: 100vh;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  .top-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 24px;
    background: #0d0d14;
    border-bottom: 1px solid var(--border);
    flex-shrink: 0;
  }
  .status {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 0.9rem;
  }
  .status-dot {
    width: 10px; height: 10px; border-radius: 50%;
    transition: background 0.3s;
  }
  .status-dot.active { background: var(--green); }
  .status-dot.waiting { background: var(--orange); animation: pulse 2s ease-in-out infinite; }
  .status-dot.keyboard { background: var(--green); }
  @keyframes pulse { 0%, 100% { opacity: 0.4; } 50% { opacity: 1; } }

  .slide-info {
    font-size: 1.1rem;
    font-weight: 600;
    color: var(--accent);
  }
  .controls {
    display: flex;
    gap: 6px;
    align-items: center;
  }
  .controls button {
    padding: 5px 12px;
    background: var(--bg-panel);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 5px;
    cursor: pointer;
    font-size: 0.8rem;
    transition: 0.15s;
  }
  .controls button:hover { background: var(--accent); border-color: var(--accent); }
  .controls span.label { font-size: 0.75rem; color: var(--dim); margin-left: 8px; }

  .timer {
    font-family: 'SF Mono', 'Fira Code', monospace;
    font-size: 1rem;
    color: var(--dim);
  }

  .progress { height: 3px; background: var(--border); flex-shrink: 0; }
  .progress .fill { height: 100%; background: var(--accent); transition: width 0.4s; width: 0; }

  .main {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  .teleprompter {
    flex: 1;
    overflow-y: auto;
    padding: 48px 60px;
    line-height: 1.9;
    font-size: 2rem;
  }
  .teleprompter .inner {
    max-width: 900px;
    margin: 0 auto;
    transition: opacity 0.12s;
  }
  .teleprompter p { margin-bottom: 1.1em; }
  .teleprompter strong { color: var(--accent); }
  .teleprompter em { color: #f0c040; }
  .teleprompter.mirror { transform: scaleX(-1); }

  .waiting-msg {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    color: var(--dim);
    gap: 16px;
    text-align: center;
  }
  .waiting-msg .icon { font-size: 2.5rem; }
  .waiting-msg p { font-size: 1rem; max-width: 500px; line-height: 1.6; }

  .footer {
    padding: 8px 24px;
    background: #0d0d14;
    border-top: 1px solid var(--border);
    display: flex;
    gap: 20px;
    font-size: 0.75rem;
    color: var(--dim);
    flex-shrink: 0;
    flex-wrap: wrap;
  }
  .footer kbd {
    background: var(--bg-panel);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 1px 5px;
    font-family: inherit;
    font-size: 0.7rem;
  }
</style>
</head>
<body>

<div class="top-bar">
  <div class="status">
    <div class="status-dot" id="statusDot"></div>
    <span id="statusText">Starting...</span>
  </div>
  <div class="slide-info" id="slideInfo">&mdash;</div>
  <div class="timer" id="timer">00:00:00</div>
  <div class="controls">
    <span class="label">Text:</span>
    <button id="fontDown">A&minus;</button>
    <button id="fontUp">A+</button>
    <label style="display:flex;align-items:center;gap:4px;font-size:0.8rem;color:var(--dim);cursor:pointer;margin-left:10px">
      <input type="checkbox" id="mirrorCheck"> Mirror
    </label>
  </div>
</div>

<div class="progress"><div class="fill" id="progressFill"></div></div>

<div class="main">
  <div class="teleprompter" id="teleprompter">
    <div class="inner" id="scriptContent">
      <div class="waiting-msg" id="waitingMsg">
        <div class="icon">&#x1F4E1;</div>
        <p>Waiting for slideshow...<br>Start your PowerPoint slideshow and the script will appear here.</p>
      </div>
    </div>
  </div>
</div>

<div class="footer" id="footer">
  <span><kbd>+</kbd><kbd>&minus;</kbd> Font size</span>
  <span><kbd>T</kbd> Timer</span>
  <span><kbd>M</kbd> Mirror</span>
  <span id="modeHint">Syncs automatically with PowerPoint via VBA</span>
</div>

<script>
let lastSlide = -1;
let fontSize = 2;
let timerRunning = false;
let timerStart = 0;
let timerElapsed = 0;
let timerInterval = null;
let wasActive = false;

async function pollState() {
  try {
    const resp = await fetch('/api/state');
    const data = await resp.json();
    const dot = document.getElementById('statusDot');
    const statusText = document.getElementById('statusText');
    const modeHint = document.getElementById('modeHint');
    const content = document.getElementById('scriptContent');
    const tp = document.getElementById('teleprompter');

    if (data.mode === 'keyboard') {
      dot.className = 'status-dot keyboard';
      statusText.textContent = 'Keyboard mode';
      modeHint.textContent = 'Period (.) = next, Comma (,) = prev';
      data.active = true;
    }

    if (data.active) {
      dot.className = 'status-dot active';
      if (data.mode !== 'keyboard') statusText.textContent = 'Slideshow active';

      document.getElementById('slideInfo').textContent =
        `Slide ${data.slide} / ${data.totalSlides}`;

      const pct = data.totalSlides > 1
        ? ((data.slide - 1) / (data.totalSlides - 1)) * 100 : 100;
      document.getElementById('progressFill').style.width = pct + '%';

      if (data.slide !== lastSlide) {
        console.log(`Slide changed: ${lastSlide} -> ${data.slide}, html length: ${(data.scriptHtml||'').length}`);
        content.innerHTML = data.scriptHtml ||
          '<p style="color:var(--dim)">No script for this slide.</p>';
        tp.scrollTo(0, 0);
        lastSlide = data.slide;
        if (!timerRunning && !timerElapsed) toggleTimer();
      }
      wasActive = true;
    } else {
      dot.className = 'status-dot waiting';
      statusText.textContent = wasActive ? 'Slideshow ended' : 'Waiting for slideshow...';
      document.getElementById('slideInfo').textContent = '\u2014';
      document.getElementById('progressFill').style.width = '0%';
      if (lastSlide !== -1) {
        content.innerHTML =
          '<div class="waiting-msg"><div class="icon">&#x1F4E1;</div>' +
          '<p>' + (wasActive ? 'Slideshow ended.' : 'Waiting for slideshow...<br>Start your PowerPoint slideshow and the script will appear here.') + '</p></div>';
        lastSlide = -1;
      }
    }
  } catch (e) { console.error('Poll error:', e); }
}

function toggleTimer() {
  if (timerRunning) {
    timerRunning = false;
    timerElapsed += Date.now() - timerStart;
    clearInterval(timerInterval);
  } else {
    timerRunning = true;
    timerStart = Date.now();
    timerInterval = setInterval(updateTimer, 1000);
    updateTimer();
  }
}

function updateTimer() {
  const total = timerElapsed + (timerRunning ? Date.now() - timerStart : 0);
  const s = Math.floor(total / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  document.getElementById('timer').textContent =
    `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(sec).padStart(2,'0')}`;
}

function changeFontSize(delta) {
  fontSize = Math.max(0.8, Math.min(5, fontSize + delta));
  document.getElementById('teleprompter').style.fontSize = fontSize + 'rem';
}

document.addEventListener('keydown', e => {
  switch (e.key) {
    case '+': case '=': changeFontSize(0.3); break;
    case '-': case '_': changeFontSize(-0.3); break;
    case 't': case 'T': toggleTimer(); break;
    case 'm': case 'M':
      const cb = document.getElementById('mirrorCheck');
      cb.checked = !cb.checked;
      document.getElementById('teleprompter').classList.toggle('mirror', cb.checked);
      break;
  }
});

document.getElementById('fontUp').addEventListener('click', () => changeFontSize(0.3));
document.getElementById('fontDown').addEventListener('click', () => changeFontSize(-0.3));
document.getElementById('mirrorCheck').addEventListener('change', e => {
  document.getElementById('teleprompter').classList.toggle('mirror', e.target.checked);
});

setInterval(pollState, 200);
pollState();
</script>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser(
        description="Slide Teleprompter — VBA event-driven"
    )
    parser.add_argument("script", help="Path to the script .md file with SLIDE markers")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"HTTP port (default: {DEFAULT_PORT})")
    parser.add_argument("--no-browser", action="store_true", help="Don't auto-open browser")
    parser.add_argument("--keyboard", action="store_true",
                        help="Use global keyboard monitoring (period/comma) instead of VBA")

    args = parser.parse_args()

    global script_sections, total_slides, mode, slideshow_active

    # Parse script
    script_path = Path(args.script)
    if not script_path.exists():
        print(f"Error: Script file not found: {args.script}")
        sys.exit(1)

    script_sections = parse_script(str(script_path))
    total_slides = len(script_sections)
    print(f"Loaded {total_slides} script sections from {script_path.name}")

    if args.keyboard:
        mode = "keyboard"
        slideshow_active = True
        start_keyboard_listener()
    else:
        mode = "vba"

        # Generate the macro with the correct port
        vba_macro = VBA_MACRO_TEMPLATE.replace("{{PORT}}", str(args.port))

        # Print the VBA macro
        print("\n" + "=" * 64)
        print("  VBA MACRO — paste this into your PowerPoint presentation")
        print("=" * 64)
        print()
        print("  1. Open your .pptx in PowerPoint")
        print("  2. Tools → Macro → Visual Basic Editor")
        print("  3. Insert → Module")
        print("  4. Paste the code below, then close the VB editor")
        print("  5. Start your slideshow — the teleprompter will sync")
        print()
        print("-" * 64)
        print(vba_macro)
        print("-" * 64)
        print()

        # Also save to a file for easy copy
        macro_path = Path(args.script).parent / "vba_macro.txt"
        macro_path.write_text(vba_macro)
        print(f"  (Also saved to {macro_path})")
        print()
        print(f"Listening for VBA callbacks on http://localhost:{args.port}/api/slide/...")

    # Start HTTP server
    server = http.server.HTTPServer(("127.0.0.1", args.port), TeleprompterHandler)
    print(f"\nTeleprompter running at http://localhost:{args.port}")
    print("Press Ctrl+C to quit.\n")

    if not args.no_browser:
        webbrowser.open(f"http://localhost:{args.port}")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        server.shutdown()


if __name__ == "__main__":
    main()
