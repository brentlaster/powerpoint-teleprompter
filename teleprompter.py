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
import platform
import re
import subprocess
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
deck_path = None  # path to .pptx/.pptm file (set via config or --deck)

# ── Display settings (controlled by remote) ──
display_settings = {
    "fontSize": 2.8,
    "textWidth": 1800,
    "wordSpacing": 0,        # px, 0 = normal
    "autoScroll": False,
    "scrollSpeed": 30,        # seconds per viewport-height
    "mirror": False,
    "settingsVersion": 0,     # bumped on every change so teleprompter knows to update
}
settings_lock = threading.Lock()

# ── Scroll commands from remote ──
# The remote posts scroll commands; the teleprompter polls and executes them.
scroll_command = {"action": None, "version": 0}  # action: "up", "down", "top"
scroll_lock = threading.Lock()

# ── Network info (set at startup) ──
local_ip = "127.0.0.1"
server_port = DEFAULT_PORT


def _activate_slideshow(slide_num):
    """Set the slideshow as active and jump to the given slide."""
    global slideshow_active, current_slide, total_slides
    slideshow_active = True
    current_slide = slide_num
    total_slides = len(script_sections)


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


# ── PowerPoint slide control via AppleScript (macOS) ──

def _ppt_applescript(action, deck_path=None):
    """Control PowerPoint via AppleScript (macOS only).

    Actions: "next", "prev", "start", "open"
    Returns (success: bool, error: str|None).
    """
    if platform.system() != "Darwin":
        return False, "PowerPoint control via remote is only supported on macOS"

    if action == "next":
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  if (count of slide show windows) > 0 then\n'
            '    go to next slide slide show view of slide show window 1\n'
            '    set idx to slide index of slide of slide show view of slide show window 1\n'
            '    set tot to count of slides of slide show settings of active presentation\n'
            '    return (idx as string) & "/" & (tot as string)\n'
            '  end if\n'
            'end tell'
        )
    elif action == "prev":
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  if (count of slide show windows) > 0 then\n'
            '    go to previous slide slide show view of slide show window 1\n'
            '    set idx to slide index of slide of slide show view of slide show window 1\n'
            '    set tot to count of slides of slide show settings of active presentation\n'
            '    return (idx as string) & "/" & (tot as string)\n'
            '  end if\n'
            'end tell'
        )
    elif action == "start":
        # Start slideshow from slide 1 on the already-open presentation
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  activate\n'
            '  if (count of presentations) > 0 then\n'
            '    set thePresentation to active presentation\n'
            '    set theSettings to slide show settings of thePresentation\n'
            '    run slide show theSettings\n'
            '  end if\n'
            'end tell'
        )
    elif action == "open" and deck_path:
        # Open a presentation file
        abs_path = str(Path(deck_path).resolve())
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  activate\n'
            f'  open "{abs_path}"\n'
            'end tell'
        )
    else:
        return False, f"Unknown action: {action}"

    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, timeout=10, text=True
        )
        if result.returncode != 0 and result.stderr.strip():
            return False, result.stderr.strip()
        return True, result.stdout.strip() if result.stdout else None
    except subprocess.TimeoutExpired:
        return False, "AppleScript timed out"
    except Exception as e:
        return False, str(e)


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

    def _json_response(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode("utf-8"))

    def do_GET(self):
        if self.path == "/":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))

        elif self.path == "/qr":
            remote_url = f"http://{local_ip}:{server_port}/remote"
            page = QR_PAGE.replace("{{REMOTE_URL}}", remote_url)
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(page.encode("utf-8"))

        elif self.path == "/api/remote-url":
            remote_url = f"http://{local_ip}:{server_port}/remote"
            self._json_response({"url": remote_url})

        elif self.path == "/remote":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(REMOTE_PAGE.encode("utf-8"))

        elif self.path == "/api/state":
            slide_idx = current_slide - 1
            section_idx = max(0, min(slide_idx, len(script_sections) - 1))
            script_text = script_sections[section_idx] if script_sections else ""

            with settings_lock:
                settings_copy = dict(display_settings)
            with scroll_lock:
                scroll_copy = dict(scroll_command)

            state = {
                "slide": current_slide,
                "totalSlides": total_slides,
                "scriptHtml": md_to_html(script_text),
                "totalSections": len(script_sections),
                "active": slideshow_active,
                "mode": mode,
                "settings": settings_copy,
                "scroll": scroll_copy,
            }
            self._json_response(state)

        elif self.path == "/api/settings":
            with settings_lock:
                self._json_response(display_settings)

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
                self._json_response({"ok": True, "slide": current_slide})
            except Exception:
                self.send_response(400)
                self.end_headers()

        elif self.path == "/api/ppt/next":
            ok, result = _ppt_applescript("next")
            if ok and result and "/" in result:
                idx, tot = result.split("/", 1)
                _activate_slideshow(int(idx))
            self._json_response({"ok": ok, "error": result if not ok else None})

        elif self.path == "/api/ppt/prev":
            ok, result = _ppt_applescript("prev")
            if ok and result and "/" in result:
                idx, tot = result.split("/", 1)
                _activate_slideshow(int(idx))
            self._json_response({"ok": ok, "error": result if not ok else None})

        elif self.path == "/api/ppt/start":
            ok, err = _ppt_applescript("start")
            if ok:
                # Activate teleprompter on slide 1 since VBA won't fire for the initial slide
                _activate_slideshow(1)
            self._json_response({"ok": ok, "error": err})

        elif self.path.startswith("/api/scroll/"):
            # Remote sends scroll commands: /api/scroll/up, /api/scroll/down, /api/scroll/top
            action = self.path.split("/")[-1]
            if action in ("up", "down", "top"):
                with scroll_lock:
                    scroll_command["action"] = action
                    scroll_command["version"] += 1
                self._json_response({"ok": True})
            else:
                self._json_response({"ok": False, "error": "unknown action"}, 400)

        elif self.path == "/api/scroll-state":
            with scroll_lock:
                self._json_response(scroll_command)

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        """Handle POST requests for settings updates from remote."""
        if self.path == "/api/settings":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(length)
                updates = json.loads(body)

                with settings_lock:
                    for key in ("fontSize", "textWidth", "wordSpacing",
                                "autoScroll", "scrollSpeed", "mirror"):
                        if key in updates:
                            display_settings[key] = updates[key]
                    display_settings["settingsVersion"] += 1
                    result = dict(display_settings)

                self._json_response(result)
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


# ── QR Code Page ──

QR_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Teleprompter Remote - QR Code</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
    background: #111118;
    color: #eaeaea;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 100vh;
    padding: 32px;
    text-align: center;
  }
  h1 { font-size: 1.6rem; margin-bottom: 8px; }
  .subtitle { color: #888; font-size: 1rem; margin-bottom: 32px; }
  #qrcode {
    background: #fff;
    padding: 24px;
    border-radius: 16px;
    display: inline-block;
    margin-bottom: 24px;
  }
  .url-display {
    font-family: 'SF Mono', 'Fira Code', monospace;
    font-size: 1.1rem;
    color: #e94560;
    background: #1a1a28;
    padding: 12px 24px;
    border-radius: 8px;
    margin-bottom: 16px;
    word-break: break-all;
  }
  .hint { color: #666; font-size: 0.85rem; margin-top: 8px; }
</style>
</head>
<body>
  <h1>Scan to Open Remote</h1>
  <p class="subtitle">Point your phone camera at this QR code</p>
  <div id="qrcode"></div>
  <div class="url-display" id="urlDisplay"></div>
  <p class="hint">Your phone must be on the same Wi-Fi network as this computer</p>

  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <script>
    // Server injects the real LAN IP — works even when opened via localhost
    var remoteUrl = '{{REMOTE_URL}}';
    document.getElementById('urlDisplay').textContent = remoteUrl;
    new QRCode(document.getElementById('qrcode'), {
      text: remoteUrl,
      width: 280,
      height: 280,
      colorDark: '#000000',
      colorLight: '#ffffff',
      correctLevel: QRCode.CorrectLevel.M
    });
  </script>
</body>
</html>
"""

# ── Remote Control Page (phone) ──

REMOTE_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
<title>Teleprompter Remote</title>
<style>
  :root {
    --bg: #111118;
    --accent: #e94560;
    --text: #eaeaea;
    --dim: #888;
    --border: #2a2a3e;
    --green: #2ecc71;
    --btn: #2a2a44;
    --slide-green: #1a6b3a;
    --scroll-blue: #2a5a8c;
    --start-orange: #b8860b;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
    background: var(--bg);
    color: var(--text);
    padding: 16px;
    touch-action: manipulation;
    -webkit-user-select: none;
    user-select: none;
    min-height: 100vh;
  }

  h1 {
    font-size: 1.1rem;
    color: var(--dim);
    text-align: center;
    margin-bottom: 6px;
    font-weight: 500;
  }
  .slide-info {
    text-align: center;
    color: var(--accent);
    font-size: 1.2rem;
    font-weight: 600;
    margin-bottom: 14px;
  }

  .btn-row {
    display: flex;
    gap: 10px;
    margin-bottom: 48px;
  }

  .big-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    padding: 20px;
    margin-bottom: 48px;
    border-radius: 14px;
    border: 2px solid var(--border);
    background: var(--btn);
    color: var(--text);
    font-size: 1.4rem;
    font-weight: 700;
    cursor: pointer;
    touch-action: manipulation;
    -webkit-tap-highlight-color: transparent;
    transition: 0.15s;
    gap: 10px;
    min-height: 80px;
  }
  .big-btn:active { transform: scale(0.97); }
  .big-btn.scrolling { background: var(--accent); border-color: var(--accent); color: #fff; }

  .half { flex: 1; margin-bottom: 0; }

  .ctrl-row {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 48px;
  }
  .ctrl-btn {
    flex: 0 0 auto;
    width: 64px;
    height: 64px;
    border-radius: 12px;
    border: 2px solid var(--border);
    background: var(--btn);
    color: var(--text);
    font-size: 1.5rem;
    font-weight: 700;
    cursor: pointer;
    touch-action: manipulation;
    -webkit-tap-highlight-color: transparent;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: 0.12s;
  }
  .ctrl-btn:active { transform: scale(0.93); background: var(--accent); border-color: var(--accent); }
  .ctrl-value {
    flex: 1;
    text-align: center;
    font-family: 'SF Mono', 'Fira Code', monospace;
    font-size: 1.1rem;
    color: var(--text);
  }
  .section-label {
    font-size: 0.85rem;
    color: var(--dim);
    margin-bottom: 6px;
    padding-left: 4px;
  }

  .status-bar {
    text-align: center;
    font-size: 0.8rem;
    color: var(--dim);
    margin-top: 10px;
  }
</style>
</head>
<body>

<h1>Teleprompter Remote</h1>
<div class="slide-info" id="slideInfo">&mdash;</div>

<!-- Start Slideshow -->
<button class="big-btn" id="startShowBtn" style="background:var(--start-orange);border-color:var(--start-orange);color:#fff;" ontouchend="startShow(event)" onclick="startShow(event)">
  &#x25B6; Start Slideshow
</button>

<!-- Slide control (side by side) -->
<div class="btn-row">
  <button class="big-btn half" style="background:var(--slide-green);border-color:var(--slide-green);" ontouchend="pptSlide('prev',event)" onclick="pptSlide('prev',event)">
    &#x25C0; Prev Slide
  </button>
  <button class="big-btn half" style="background:var(--slide-green);border-color:var(--slide-green);" ontouchend="pptSlide('next',event)" onclick="pptSlide('next',event)">
    Next Slide &#x25B6;
  </button>
</div>

<!-- Scroll teleprompter text (side by side) -->
<div class="btn-row">
  <button class="big-btn half" style="background:var(--scroll-blue);border-color:var(--scroll-blue);" ontouchend="scrollText('up',event)" onclick="scrollText('up',event)">
    &#x25B2; Scroll Up
  </button>
  <button class="big-btn half" style="background:var(--scroll-blue);border-color:var(--scroll-blue);" ontouchend="scrollText('down',event)" onclick="scrollText('down',event)">
    Scroll Down &#x25BC;
  </button>
</div>

<!-- Auto-scroll toggle -->
<button class="big-btn" id="autoScrollBtn" ontouchend="toggleScroll(event)" onclick="toggleScroll(event)">
  &#x25B6; Start Auto-Scroll
</button>

<!-- Scroll speed -->
<div class="section-label">Scroll Speed</div>
<div class="ctrl-row">
  <button class="ctrl-btn" ontouchend="adjSpeed(-5,event)" onclick="adjSpeed(-5,event)">&#x2212;</button>
  <span class="ctrl-value" id="speedVal">30s</span>
  <button class="ctrl-btn" ontouchend="adjSpeed(5,event)" onclick="adjSpeed(5,event)">+</button>
</div>

<!-- Font size -->
<div class="section-label">Font Size</div>
<div class="ctrl-row">
  <button class="ctrl-btn" style="font-size:1.1rem;" ontouchend="adjFont(-0.3,event)" onclick="adjFont(-0.3,event)">A&#x2212;</button>
  <span class="ctrl-value" id="fontVal">2.8</span>
  <button class="ctrl-btn" style="font-size:1.1rem;" ontouchend="adjFont(0.3,event)" onclick="adjFont(0.3,event)">A+</button>
</div>

<!-- Text width -->
<div class="section-label">Text Width</div>
<div class="ctrl-row">
  <button class="ctrl-btn" ontouchend="adjWidth(-200,event)" onclick="adjWidth(-200,event)">&#x2190;</button>
  <span class="ctrl-value" id="widthVal">1800</span>
  <button class="ctrl-btn" ontouchend="adjWidth(200,event)" onclick="adjWidth(200,event)">&#x2192;</button>
</div>

<!-- Word spacing -->
<div class="section-label">Word Spacing</div>
<div class="ctrl-row">
  <button class="ctrl-btn" ontouchend="adjWordSp(-2,event)" onclick="adjWordSp(-2,event)">&#x2212;</button>
  <span class="ctrl-value" id="wordSpVal">0px</span>
  <button class="ctrl-btn" ontouchend="adjWordSp(2,event)" onclick="adjWordSp(2,event)">+</button>
</div>

<div class="status-bar" id="statusBar">Connecting...</div>

<script>
var settings = {};
var scrolling = false;

function stopEvent(e) {
  if (e) { e.preventDefault(); e.stopPropagation(); }
}

function startShow(e) {
  stopEvent(e);
  fetch('/api/ppt/start')
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (!data.ok && data.error) {
        document.getElementById('statusBar').textContent = 'Error: ' + data.error;
      } else {
        document.getElementById('statusBar').textContent = 'Slideshow starting...';
      }
    })
    .catch(function() {
      document.getElementById('statusBar').textContent = 'Connection error';
    });
}

function pptSlide(dir, e) {
  stopEvent(e);
  fetch('/api/ppt/' + dir)
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (!data.ok && data.error) {
        document.getElementById('statusBar').textContent = 'Error: ' + data.error;
      }
    })
    .catch(function() {
      document.getElementById('statusBar').textContent = 'Connection error';
    });
}

function scrollText(dir, e) {
  stopEvent(e);
  fetch('/api/scroll/' + dir)
    .catch(function() {
      document.getElementById('statusBar').textContent = 'Connection error';
    });
}

function postSettings(updates) {
  fetch('/api/settings', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(updates)
  }).then(function(r) { return r.json(); })
    .then(function(s) { settings = s; updateUI(); });
}

function toggleScroll(e) {
  stopEvent(e);
  scrolling = !scrolling;
  postSettings({ autoScroll: scrolling });
}

function adjSpeed(d, e) {
  stopEvent(e);
  var v = Math.max(5, Math.min(120, (settings.scrollSpeed || 30) + d));
  postSettings({ scrollSpeed: v });
}

function adjFont(d, e) {
  stopEvent(e);
  var v = Math.max(1.0, Math.min(6.0, (settings.fontSize || 2.8) + d));
  v = Math.round(v * 10) / 10;
  postSettings({ fontSize: v });
}

function adjWidth(d, e) {
  stopEvent(e);
  var v = Math.max(600, Math.min(3000, (settings.textWidth || 1800) + d));
  postSettings({ textWidth: v });
}

function adjWordSp(d, e) {
  stopEvent(e);
  var v = Math.max(-4, Math.min(24, (settings.wordSpacing || 0) + d));
  postSettings({ wordSpacing: v });
}

function updateUI() {
  scrolling = settings.autoScroll || false;
  var btn = document.getElementById('autoScrollBtn');
  if (scrolling) {
    btn.className = 'big-btn scrolling';
    btn.innerHTML = '&#x23F8; Pause Auto-Scroll';
  } else {
    btn.className = 'big-btn';
    btn.innerHTML = '&#x25B6; Start Auto-Scroll';
  }
  document.getElementById('speedVal').textContent = (settings.scrollSpeed || 30) + 's';
  document.getElementById('fontVal').textContent = (settings.fontSize || 2.8).toFixed(1);
  document.getElementById('widthVal').textContent = (settings.textWidth || 1800);
  document.getElementById('wordSpVal').textContent = (settings.wordSpacing || 0) + 'px';
}

function poll() {
  fetch('/api/state')
    .then(function(r) { return r.json(); })
    .then(function(data) {
      var info = document.getElementById('slideInfo');
      var startBtn = document.getElementById('startShowBtn');
      if (data.active) {
        info.textContent = 'Slide ' + data.slide + ' / ' + data.totalSlides;
        startBtn.style.display = 'none';
      } else {
        info.textContent = 'Waiting for slideshow...';
        startBtn.style.display = 'flex';
      }
      if (data.settings) {
        settings = data.settings;
        updateUI();
      }
      document.getElementById('statusBar').textContent = 'Connected';
    })
    .catch(function() {
      document.getElementById('statusBar').textContent = 'Connection lost - retrying...';
    });
}

setInterval(poll, 500);
poll();

// ── Keep screen awake (Wake Lock API) ──
var wakeLock = null;
async function requestWakeLock() {
  try {
    if ('wakeLock' in navigator) {
      wakeLock = await navigator.wakeLock.request('screen');
      wakeLock.addEventListener('release', function() { wakeLock = null; });
    }
  } catch (e) {}
}
requestWakeLock();
// Re-acquire wake lock when page becomes visible again (e.g. after tab switch)
document.addEventListener('visibilitychange', function() {
  if (document.visibilityState === 'visible') requestWakeLock();
});
</script>
</body>
</html>
"""

# ── HTML Page ──

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
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
    --btn-bg: #2a2a44;
    --btn-hover: #e94560;
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
    touch-action: manipulation;
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
    position: relative;
  }

  .teleprompter {
    flex: 1;
    overflow-y: auto;
    padding: 48px 40px;
    line-height: 1.9;
    font-size: 2.8rem;
  }
  .teleprompter .inner {
    max-width: 1800px;
    margin: 0 auto;
    transition: opacity 0.12s;
  }
  .teleprompter p { margin-bottom: 1.1em; }
  .teleprompter strong { color: var(--accent); }
  .teleprompter em { color: #f0c040; }
  .teleprompter.mirror { transform: scaleX(-1); }

  /* Auto-scroll progress bar at top of text area */
  .scroll-progress {
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 4px;
    background: var(--border);
    z-index: 5;
    display: none;
  }
  .scroll-progress .scroll-fill {
    height: 100%;
    background: var(--green);
    transition: width 0.3s linear;
    width: 0;
  }
  .scroll-progress.active { display: block; }

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

  /* ── Floating control panel (bottom-right) ── */
  .control-panel {
    position: absolute;
    bottom: 20px;
    right: 20px;
    background: rgba(13, 13, 20, 0.95);
    border: 2px solid var(--border);
    border-radius: 16px;
    padding: 0;
    z-index: 20;
    backdrop-filter: blur(8px);
    min-width: 340px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.5);
    transition: opacity 0.2s;
    touch-action: manipulation;
    -webkit-user-select: none;
    user-select: none;
  }
  .control-panel.collapsed .cp-body { display: none; }
  .control-panel.collapsed { min-width: auto; }

  .cp-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 18px;
    cursor: pointer;
    user-select: none;
    border-bottom: 2px solid var(--border);
    font-size: 1.05rem;
    font-weight: 600;
    color: var(--dim);
    touch-action: manipulation;
    min-height: 54px;
  }
  .control-panel.collapsed .cp-header { border-bottom: none; }
  .cp-header:hover { color: var(--text); }
  .cp-toggle { font-size: 0.85rem; color: var(--dim); }

  .cp-body { padding: 16px 18px; }

  .cp-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 14px;
    gap: 10px;
  }
  .cp-row:last-child { margin-bottom: 0; }

  .cp-label {
    font-size: 1rem;
    color: var(--dim);
    white-space: nowrap;
    min-width: 70px;
  }

  .cp-btn-group {
    display: flex;
    gap: 6px;
    align-items: center;
  }

  .cp-btn {
    padding: 12px 18px;
    background: var(--btn-bg);
    color: var(--text);
    border: 2px solid var(--border);
    border-radius: 10px;
    cursor: pointer;
    font-size: 1.3rem;
    font-weight: 700;
    transition: 0.15s;
    touch-action: manipulation;
    -webkit-tap-highlight-color: transparent;
    min-width: 60px;
    min-height: 60px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .cp-btn:hover { background: var(--btn-hover); border-color: var(--btn-hover); }
  .cp-btn:active { transform: scale(0.93); background: var(--btn-hover); border-color: var(--btn-hover); }
  .cp-btn.active-btn { background: var(--green); border-color: var(--green); color: #111; }
  .cp-btn.stop-btn { background: var(--accent); border-color: var(--accent); }

  .cp-value {
    font-family: 'SF Mono', 'Fira Code', monospace;
    font-size: 1.05rem;
    color: var(--text);
    min-width: 50px;
    text-align: center;
  }

  /* Auto-scroll big toggle */
  .cp-start-btn {
    width: 100%;
    padding: 12px 8px;
    margin-bottom: 10px;
    border-radius: 10px;
    border: 2px solid #b8860b;
    background: #b8860b;
    color: #fff;
    font-size: 0.95rem;
    font-weight: 700;
    cursor: pointer;
    touch-action: manipulation;
    -webkit-tap-highlight-color: transparent;
    transition: 0.12s;
    min-height: 48px;
  }
  .cp-start-btn:hover { background: #d4a017; border-color: #d4a017; }
  .cp-start-btn:active { transform: scale(0.95); }

  .cp-slide-row {
    display: flex;
    gap: 8px;
    margin-bottom: 12px;
  }
  .cp-slide-btn {
    flex: 1;
    padding: 12px 8px;
    border-radius: 10px;
    border: 2px solid #1a6b3a;
    background: #1a6b3a;
    color: #fff;
    font-size: 0.95rem;
    font-weight: 700;
    cursor: pointer;
    touch-action: manipulation;
    -webkit-tap-highlight-color: transparent;
    transition: 0.12s;
    min-height: 48px;
  }
  .cp-slide-btn:hover { background: #238c4e; border-color: #238c4e; }
  .cp-slide-btn:active { transform: scale(0.95); }

  .autoscroll-toggle {
    width: 100%;
    padding: 14px;
    margin-bottom: 14px;
    border-radius: 10px;
    border: 2px solid var(--border);
    background: var(--btn-bg);
    color: var(--text);
    font-size: 1.2rem;
    font-weight: 700;
    cursor: pointer;
    touch-action: manipulation;
    -webkit-tap-highlight-color: transparent;
    transition: 0.15s;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
    min-height: 64px;
  }
  .autoscroll-toggle:hover { background: var(--green); color: #111; border-color: var(--green); }
  .autoscroll-toggle.scrolling { background: var(--accent); border-color: var(--accent); color: #fff; }
  .autoscroll-toggle.scrolling:hover { background: #c0374d; }

  /* Touch overlay for pause/resume */
  .touch-hint {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    background: rgba(0,0,0,0.7);
    color: #fff;
    padding: 16px 32px;
    border-radius: 12px;
    font-size: 1.3rem;
    font-weight: 600;
    pointer-events: none;
    opacity: 0;
    transition: opacity 0.3s;
    z-index: 30;
  }
  .touch-hint.show { opacity: 1; }

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
  /* QR code in control panel */
  .cp-qr-row {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 8px 0 10px 0;
    border-bottom: 1px solid var(--border);
    margin-bottom: 8px;
  }
  .cp-qr-row canvas, .cp-qr-row img {
    border-radius: 6px;
    width: 160px !important;
    height: 160px !important;
    min-width: 160px;
    min-height: 160px;
    display: block;
  }
  #cpQrCode {
    flex-shrink: 0;
    width: 160px;
    height: 160px;
  }
  .cp-qr-label {
    font-size: 0.75rem;
    color: var(--dim);
    line-height: 1.3;
  }
  .cp-qr-label a { color: var(--accent); text-decoration: none; }
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
</head>
<body>

<div class="top-bar">
  <div class="status">
    <div class="status-dot" id="statusDot"></div>
    <span id="statusText">Starting...</span>
  </div>
  <div class="slide-info" id="slideInfo">&mdash;</div>
  <div class="timer" id="timer">00:00:00</div>
</div>

<div class="progress"><div class="fill" id="progressFill"></div></div>

<div class="main">
  <div class="scroll-progress" id="scrollProgress">
    <div class="scroll-fill" id="scrollFill"></div>
  </div>

  <div class="teleprompter" id="teleprompter">
    <div class="inner" id="scriptContent">
      <div class="waiting-msg" id="waitingMsg">
        <div class="icon">&#x1F4E1;</div>
        <p>Waiting for slideshow...<br>Start your PowerPoint slideshow and the script will appear here.</p>
      </div>
    </div>
  </div>

  <div class="touch-hint" id="touchHint"></div>

  <!-- Floating control panel -->
  <div class="control-panel" id="controlPanel">
    <div class="cp-header" id="cpHeader">
      <span>Controls</span>
      <span class="cp-toggle" id="cpToggle">&#x25BC;</span>
    </div>
    <div class="cp-body" id="cpBody">

      <!-- QR code for phone remote -->
      <div class="cp-qr-row">
        <div id="cpQrCode"></div>
        <div class="cp-qr-label">Scan for<br><a href="#" id="cpQrLink">phone remote</a></div>
      </div>

      <!-- Start slideshow -->
      <button class="cp-start-btn" id="cpStartShow">&#x25B6; Start Slideshow</button>

      <!-- Slide control -->
      <div class="cp-slide-row">
        <button class="cp-slide-btn" id="cpPrevSlide">&#x25C0; Prev Slide</button>
        <button class="cp-slide-btn" id="cpNextSlide">Next Slide &#x25B6;</button>
      </div>

      <!-- Auto-Scroll toggle -->
      <button class="autoscroll-toggle" id="autoScrollBtn">
        &#x25B6; Start Auto-Scroll
      </button>

      <!-- Scroll speed -->
      <div class="cp-row">
        <span class="cp-label">Speed</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="speedDown">&#x2212;</button>
          <span class="cp-value" id="speedValue">30s</span>
          <button class="cp-btn" id="speedUp">+</button>
        </div>
      </div>

      <!-- Font size -->
      <div class="cp-row">
        <span class="cp-label">Font</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="fontDown" style="font-size:1.1rem;">A&#x2212;</button>
          <span class="cp-value" id="fontValue">2.8</span>
          <button class="cp-btn" id="fontUp" style="font-size:1.1rem;">A+</button>
        </div>
      </div>

      <!-- Text width -->
      <div class="cp-row">
        <span class="cp-label">Width</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="widthDown">&#x2190;</button>
          <span class="cp-value" id="widthValue">1800</span>
          <button class="cp-btn" id="widthUp">&#x2192;</button>
        </div>
      </div>

      <!-- Word spacing -->
      <div class="cp-row">
        <span class="cp-label">Spacing</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="wordSpDown">&#x2212;</button>
          <span class="cp-value" id="wordSpValue">0px</span>
          <button class="cp-btn" id="wordSpUp">+</button>
        </div>
      </div>

      <!-- Mirror -->
      <div class="cp-row">
        <span class="cp-label">Mirror</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="mirrorBtn">Off</button>
        </div>
      </div>

    </div>
  </div>
</div>

<div class="footer" id="footer">
  <span><kbd>S</kbd> Auto-scroll</span>
  <span><kbd>T</kbd> Timer</span>
  <span>Open <strong>/remote</strong> on your phone to control without touching this screen</span>
  <span id="modeHint">Syncs via VBA</span>
</div>

<script>
var lastSlide = -1;
var fontSize = 2.8;
var timerRunning = false;
var timerStart = 0;
var timerElapsed = 0;
var timerInterval = null;
var wasActive = false;

// ── Remote scroll commands ──
var lastScrollVersion = 0;
var SCROLL_AMOUNT = 300;  // pixels per button press

function applyScrollCommand(scrollData) {
  if (!scrollData || scrollData.version === lastScrollVersion) return;
  lastScrollVersion = scrollData.version;
  var tp = document.getElementById('teleprompter');
  if (scrollData.action === 'down') {
    tp.scrollBy({ top: SCROLL_AMOUNT, behavior: 'smooth' });
  } else if (scrollData.action === 'up') {
    tp.scrollBy({ top: -SCROLL_AMOUNT, behavior: 'smooth' });
  } else if (scrollData.action === 'top') {
    tp.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

// ── Text width ──
var textWidth = 1800;
var WIDTH_STEP = 200, WIDTH_MIN = 600, WIDTH_MAX = 3000;

// ── Word spacing ──
var wordSpacing = 0;  // px
var WORDSP_STEP = 2, WORDSP_MIN = -4, WORDSP_MAX = 24;

// ── Auto-scroll ──
var autoScrolling = false;
var scrollSpeed = 30;
var SPEED_MIN = 5, SPEED_MAX = 120, SPEED_STEP = 5;
var scrollAnimFrame = null;
var lastScrollTime = 0;

// ── Settings sync with server (for remote control) ──
var lastSettingsVersion = 0;

function pushSettings() {
  // Push current local state to server so remote stays in sync
  fetch('/api/settings', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      fontSize: fontSize,
      textWidth: textWidth,
      wordSpacing: wordSpacing,
      autoScroll: autoScrolling,
      scrollSpeed: scrollSpeed,
      mirror: mirrorOn
    })
  }).catch(function() {});
}

function applyServerSettings(s) {
  if (!s || s.settingsVersion === lastSettingsVersion) return;
  lastSettingsVersion = s.settingsVersion;

  // Apply each setting if it differs
  if (s.fontSize !== fontSize) {
    fontSize = s.fontSize;
    document.getElementById('teleprompter').style.fontSize = fontSize + 'rem';
    document.getElementById('fontValue').textContent = fontSize.toFixed(1);
  }
  if (s.textWidth !== textWidth) {
    textWidth = s.textWidth;
    updateWidthDisplay();
  }
  if (s.wordSpacing !== wordSpacing) {
    wordSpacing = s.wordSpacing;
    applyWordSpacing();
  }
  if (s.scrollSpeed !== scrollSpeed) {
    scrollSpeed = s.scrollSpeed;
    updateSpeedDisplay();
  }
  if (s.mirror !== mirrorOn) {
    mirrorOn = s.mirror;
    document.getElementById('teleprompter').classList.toggle('mirror', mirrorOn);
    document.getElementById('mirrorBtn').textContent = mirrorOn ? 'On' : 'Off';
  }
  // Auto-scroll: start/stop based on server state
  if (s.autoScroll && !autoScrolling) {
    startAutoScroll(true);  // true = silent (no push back)
  } else if (!s.autoScroll && autoScrolling) {
    stopAutoScroll(true);
  }
}

// ── Display helpers ──

function updateWidthDisplay() {
  document.querySelector('.teleprompter .inner').style.maxWidth = textWidth + 'px';
  document.getElementById('widthValue').textContent = textWidth;
}

function applyWordSpacing() {
  document.querySelector('.teleprompter .inner').style.wordSpacing = wordSpacing + 'px';
  document.getElementById('wordSpValue').textContent = wordSpacing + 'px';
}

function updateSpeedDisplay() {
  document.getElementById('speedValue').textContent = scrollSpeed + 's';
}

function changeWidth(delta) {
  textWidth = Math.max(WIDTH_MIN, Math.min(WIDTH_MAX, textWidth + delta));
  updateWidthDisplay();
  pushSettings();
}

function changeWordSpacing(delta) {
  wordSpacing = Math.max(WORDSP_MIN, Math.min(WORDSP_MAX, wordSpacing + delta));
  applyWordSpacing();
  pushSettings();
}

function changeSpeed(delta) {
  scrollSpeed = Math.max(SPEED_MIN, Math.min(SPEED_MAX, scrollSpeed + delta));
  updateSpeedDisplay();
  pushSettings();
}

function startAutoScroll(silent) {
  if (autoScrolling) return;
  autoScrolling = true;
  var btn = document.getElementById('autoScrollBtn');
  btn.classList.add('scrolling');
  btn.innerHTML = '&#x23F8; Pause Auto-Scroll';
  document.getElementById('scrollProgress').classList.add('active');
  lastScrollTime = performance.now();
  scrollAnimFrame = requestAnimationFrame(doAutoScroll);
  if (!silent) {
    showTouchHint('Auto-scroll started');
    pushSettings();
  }
}

function stopAutoScroll(silent) {
  if (!autoScrolling) return;
  autoScrolling = false;
  var btn = document.getElementById('autoScrollBtn');
  btn.classList.remove('scrolling');
  btn.innerHTML = '&#x25B6; Resume Auto-Scroll';
  document.getElementById('scrollProgress').classList.remove('active');
  if (scrollAnimFrame) cancelAnimationFrame(scrollAnimFrame);
  if (!silent) {
    showTouchHint('Auto-scroll paused');
    pushSettings();
  }
}

function toggleAutoScroll() {
  if (autoScrolling) stopAutoScroll();
  else startAutoScroll();
}

function doAutoScroll(now) {
  if (!autoScrolling) return;
  var tp = document.getElementById('teleprompter');
  var dt = (now - lastScrollTime) / 1000;
  lastScrollTime = now;

  var pxPerSec = tp.clientHeight / scrollSpeed;
  tp.scrollTop += pxPerSec * dt;

  var maxScroll = tp.scrollHeight - tp.clientHeight;
  if (maxScroll > 0) {
    var pct = (tp.scrollTop / maxScroll) * 100;
    document.getElementById('scrollFill').style.width = pct + '%';
  }

  if (tp.scrollTop >= tp.scrollHeight - tp.clientHeight - 2) {
    stopAutoScroll();
    document.getElementById('autoScrollBtn').innerHTML = '&#x25B6; Start Auto-Scroll';
    showTouchHint('Reached end of text');
    return;
  }

  scrollAnimFrame = requestAnimationFrame(doAutoScroll);
}

// ── Touch hint overlay ──
var touchHintTimer = null;
function showTouchHint(msg) {
  var hint = document.getElementById('touchHint');
  hint.textContent = msg;
  hint.classList.add('show');
  clearTimeout(touchHintTimer);
  touchHintTimer = setTimeout(function() { hint.classList.remove('show'); }, 1200);
}

// ── Touch / click to pause/resume in text area ──
var touchMoved = false;
document.getElementById('teleprompter').addEventListener('touchstart', function(e) {
  if (e.target.closest('.control-panel')) return;
  touchMoved = false;
}, { passive: true });

document.getElementById('teleprompter').addEventListener('touchmove', function() {
  touchMoved = true;
}, { passive: true });

document.getElementById('teleprompter').addEventListener('touchend', function(e) {
  if (e.target.closest('.control-panel')) return;
  if (touchMoved) return;
  if (autoScrolling) { stopAutoScroll(); }
  else if (document.getElementById('autoScrollBtn').innerHTML.indexOf('Resume') !== -1) { startAutoScroll(); }
}, { passive: true });

document.getElementById('teleprompter').addEventListener('click', function(e) {
  if (e.target.closest('.control-panel')) return;
  if (autoScrolling) { stopAutoScroll(); }
  else if (document.getElementById('autoScrollBtn').innerHTML.indexOf('Resume') !== -1) { startAutoScroll(); }
});

// ── Polling (also syncs server settings from remote) ──
async function pollState() {
  try {
    var resp = await fetch('/api/state');
    var data = await resp.json();
    var dot = document.getElementById('statusDot');
    var statusText = document.getElementById('statusText');
    var modeHint = document.getElementById('modeHint');
    var content = document.getElementById('scriptContent');
    var tp = document.getElementById('teleprompter');

    // Apply any settings changes from remote
    if (data.settings) applyServerSettings(data.settings);
    // Apply any scroll commands from remote
    if (data.scroll) applyScrollCommand(data.scroll);

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
        'Slide ' + data.slide + ' / ' + data.totalSlides;

      var pct = data.totalSlides > 1
        ? ((data.slide - 1) / (data.totalSlides - 1)) * 100 : 100;
      document.getElementById('progressFill').style.width = pct + '%';

      if (data.slide !== lastSlide) {
        content.innerHTML = data.scriptHtml ||
          '<p style="color:var(--dim)">No script for this slide.</p>';
        tp.scrollTo(0, 0);
        lastSlide = data.slide;
        if (autoScrolling) {
          stopAutoScroll(true);
          setTimeout(function() { startAutoScroll(); }, 300);
        }
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
        if (autoScrolling) stopAutoScroll(true);
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
  var total = timerElapsed + (timerRunning ? Date.now() - timerStart : 0);
  var s = Math.floor(total / 1000);
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var sec = s % 60;
  document.getElementById('timer').textContent =
    String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0') + ':' + String(sec).padStart(2,'0');
}

function changeFontSize(delta) {
  fontSize = Math.max(0.8, Math.min(6, fontSize + delta));
  fontSize = Math.round(fontSize * 10) / 10;
  document.getElementById('teleprompter').style.fontSize = fontSize + 'rem';
  document.getElementById('fontValue').textContent = fontSize.toFixed(1);
  pushSettings();
}

var mirrorOn = false;
function toggleMirror() {
  mirrorOn = !mirrorOn;
  document.getElementById('teleprompter').classList.toggle('mirror', mirrorOn);
  document.getElementById('mirrorBtn').textContent = mirrorOn ? 'On' : 'Off';
  pushSettings();
}

// ── Keyboard shortcuts (no Space — conflicts with PowerPoint) ──
document.addEventListener('keydown', function(e) {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
  switch (e.key) {
    case '+': case '=': changeFontSize(0.3); break;
    case '-': case '_': changeFontSize(-0.3); break;
    case 't': case 'T': toggleTimer(); break;
    case 'm': case 'M': toggleMirror(); break;
    case 's': case 'S': toggleAutoScroll(); break;
  }
});

// ── Touch-safe button wiring ──
function addTouchButton(id, handler) {
  var el = document.getElementById(id);
  el.addEventListener('touchend', function(e) {
    e.preventDefault();
    e.stopPropagation();
    handler();
  });
  el.addEventListener('click', function(e) {
    e.stopPropagation();
    handler();
  });
}

addTouchButton('cpStartShow', function() {
  fetch('/api/ppt/start').catch(function(){});
});
addTouchButton('cpPrevSlide', function() {
  fetch('/api/ppt/prev').catch(function(){});
});
addTouchButton('cpNextSlide', function() {
  fetch('/api/ppt/next').catch(function(){});
});
addTouchButton('fontUp', function() { changeFontSize(0.3); });
addTouchButton('fontDown', function() { changeFontSize(-0.3); });
addTouchButton('widthUp', function() { changeWidth(WIDTH_STEP); });
addTouchButton('widthDown', function() { changeWidth(-WIDTH_STEP); });
addTouchButton('speedUp', function() { changeSpeed(SPEED_STEP); });
addTouchButton('speedDown', function() { changeSpeed(-SPEED_STEP); });
addTouchButton('wordSpUp', function() { changeWordSpacing(WORDSP_STEP); });
addTouchButton('wordSpDown', function() { changeWordSpacing(-WORDSP_STEP); });
addTouchButton('mirrorBtn', toggleMirror);
addTouchButton('autoScrollBtn', toggleAutoScroll);

// ── Collapsible panel (touch-safe) ──
(function() {
  var hdr = document.getElementById('cpHeader');
  function doToggle(e) {
    e.preventDefault();
    e.stopPropagation();
    var panel = document.getElementById('controlPanel');
    panel.classList.toggle('collapsed');
    document.getElementById('cpToggle').innerHTML =
      panel.classList.contains('collapsed') ? '&#x25B2;' : '&#x25BC;';
  }
  hdr.addEventListener('touchend', doToggle);
  hdr.addEventListener('click', doToggle);
})();

// ── Init ──
updateWidthDisplay();
applyWordSpacing();
updateSpeedDisplay();
document.getElementById('fontValue').textContent = fontSize.toFixed(1);

// ── QR code in control panel ──
fetch('/api/remote-url')
  .then(function(r) { return r.json(); })
  .then(function(d) {
    var link = document.getElementById('cpQrLink');
    link.href = d.url;
    link.textContent = 'phone remote';
    new QRCode(document.getElementById('cpQrCode'), {
      text: d.url,
      width: 160,
      height: 160,
      colorDark: '#e94560',
      colorLight: '#111118',
      correctLevel: QRCode.CorrectLevel.M
    });
  });

setInterval(pollState, 200);
pollState();
</script>
</body>
</html>
"""


def _detect_local_ip():
    """Detect the LAN IP address for phone remote access."""
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return None


def _load_config(config_path):
    """Load a JSON config file. Returns dict with keys:
        script, deck, port, keyboard, auto_open_deck, auto_start_show
    """
    p = Path(config_path).resolve()
    if not p.exists():
        print(f"Error: Config file not found: {config_path}")
        sys.exit(1)

    with open(p, "r") as f:
        raw = json.load(f)

    # Resolve paths relative to config file location
    config_dir = p.parent

    cfg = {}
    if "script" in raw:
        sp = Path(raw["script"])
        cfg["script"] = str(sp if sp.is_absolute() else config_dir / sp)
    if "deck" in raw:
        dp = Path(raw["deck"])
        cfg["deck"] = str(dp if dp.is_absolute() else config_dir / dp)

    cfg["port"] = raw.get("port", DEFAULT_PORT)
    cfg["keyboard"] = raw.get("keyboard", False)
    cfg["auto_open_deck"] = raw.get("auto_open_deck", True)
    # Accept both "auto_start" and "auto_start_show"
    cfg["auto_start_show"] = raw.get("auto_start_show", raw.get("auto_start", False))
    return cfg


def main():
    parser = argparse.ArgumentParser(
        description="Slide Teleprompter — VBA event-driven"
    )
    parser.add_argument("script", nargs="?", default=None,
                        help="Path to the script .md file with SLIDE markers")
    parser.add_argument("--config", "-c", help="Path to a JSON config file")
    parser.add_argument("--deck", help="Path to .pptx/.pptm file (opens in PowerPoint)")
    parser.add_argument("--port", type=int, default=None,
                        help=f"HTTP port (default: {DEFAULT_PORT})")
    parser.add_argument("--no-browser", action="store_true",
                        help="Don't auto-open browser")
    parser.add_argument("--keyboard", action="store_true",
                        help="Use global keyboard monitoring (period/comma) instead of VBA")

    args = parser.parse_args()

    global script_sections, total_slides, mode, slideshow_active
    global local_ip, server_port, deck_path

    # ── Merge config file + command-line args ──
    cfg = {}
    if args.config:
        cfg = _load_config(args.config)

    # Command-line overrides config
    script_file = args.script or cfg.get("script")
    if not script_file:
        parser.error("A script file is required (positional arg or 'script' in config)")

    deck_path = args.deck or cfg.get("deck")
    port = args.port or cfg.get("port", DEFAULT_PORT)
    use_keyboard = args.keyboard or cfg.get("keyboard", False)
    auto_open_deck = cfg.get("auto_open_deck", True) if deck_path else False
    auto_start_show = cfg.get("auto_start_show", False)

    server_port = port

    # ── Parse script ──
    script_path = Path(script_file)
    if not script_path.exists():
        print(f"Error: Script file not found: {script_file}")
        sys.exit(1)

    script_sections = parse_script(str(script_path))
    total_slides = len(script_sections)
    print(f"Loaded {total_slides} script sections from {script_path.name}")

    if deck_path:
        print(f"Deck: {deck_path}")

    # ── Mode setup ──
    if use_keyboard:
        mode = "keyboard"
        slideshow_active = True
        start_keyboard_listener()
    else:
        mode = "vba"

        vba_macro = VBA_MACRO_TEMPLATE.replace("{{PORT}}", str(port))

        if not args.config:
            # Only print VBA instructions when not using config (first-time setup)
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

        macro_path = Path(script_file).parent / "vba_macro.txt"
        macro_path.write_text(vba_macro)
        if not args.config:
            print(f"  (Also saved to {macro_path})")
            print()
        print(f"Listening for VBA callbacks on http://localhost:{port}/api/slide/...")

    # ── Auto-open deck in PowerPoint ──
    if deck_path and auto_open_deck and platform.system() == "Darwin":
        dp = Path(deck_path)
        if dp.exists():
            print(f"Opening {dp.name} in PowerPoint...")
            ok, err = _ppt_applescript("open", deck_path=str(dp))
            if not ok:
                print(f"  Warning: Could not open deck: {err}")
        else:
            print(f"  Warning: Deck file not found: {deck_path}")

    # ── Start HTTP server ──
    server = http.server.HTTPServer(("0.0.0.0", port), TeleprompterHandler)

    local_ip = _detect_local_ip() or "YOUR_IP"
    if local_ip == "YOUR_IP":
        print("  Warning: Could not detect LAN IP. Replace YOUR_IP with your computer's IP.")

    print(f"\nTeleprompter running at http://localhost:{port}")
    print(f"\n  Phone remote:  http://{local_ip}:{port}/remote")
    print(f"  QR code page:  http://localhost:{port}/qr")
    print("\nPress Ctrl+C to quit.\n")

    # ── Auto-open browser (in background so server is ready first) ──
    if not args.no_browser:
        def _open_browser():
            import time as _time
            _time.sleep(1)
            webbrowser.open(f"http://localhost:{port}")
        threading.Thread(target=_open_browser, daemon=True).start()

    # ── Auto-start slideshow (after brief delay for PowerPoint to load) ──
    if auto_start_show and deck_path and platform.system() == "Darwin":
        def _delayed_start():
            import time as _time
            _time.sleep(4)  # give PowerPoint time to open the deck
            print("Auto-starting slideshow...")
            ok, err = _ppt_applescript("start")
            if not ok:
                print(f"  Warning: Could not start slideshow: {err}")
        t = threading.Thread(target=_delayed_start, daemon=True)
        t.start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        server.shutdown()


if __name__ == "__main__":
    main()
