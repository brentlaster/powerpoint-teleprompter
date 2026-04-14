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
slide_titles = []
slideshow_active = False
mode = "vba"  # "vba" or "keyboard"
deck_path = None  # path to .pptx/.pptm file (set via config or --deck)
live_screenshot_path = None  # path to latest screenshot of slideshow window
live_screenshot_lock = threading.Lock()

# ── Demo mode state ──
demo_mode = False
demo_script = None       # path to demo-*.py (auto-detected from script directory)
demo_dir = None          # directory containing the demo script
demo_slide = None        # slide number when demo was entered (to resume from)

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

# Path to persistent settings file (set at startup based on script location)
_settings_file_path = None

# Keys that we persist to disk (excludes volatile ones like settingsVersion)
_PERSIST_KEYS = ("fontSize", "textWidth", "wordSpacing", "mirror",
                 "highlightLevel", "highlightLines", "screenshotDisplay",
                 "panelHeightPct", "uiScaleLevel")


def _load_saved_settings():
    """Load settings from the persistent JSON file, if it exists."""
    if not _settings_file_path:
        return
    try:
        with open(_settings_file_path, "r") as f:
            saved = json.load(f)
        with settings_lock:
            for key in _PERSIST_KEYS:
                if key in saved:
                    display_settings[key] = saved[key]
        print(f"Loaded saved settings from {_settings_file_path}")
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"Warning: Could not load settings file: {e}")


def _save_settings():
    """Save current settings to the persistent JSON file."""
    if not _settings_file_path:
        return
    try:
        with settings_lock:
            to_save = {k: display_settings[k] for k in _PERSIST_KEYS
                       if k in display_settings}
        with open(_settings_file_path, "w") as f:
            json.dump(to_save, f, indent=2)
    except Exception as e:
        print(f"Warning: Could not save settings file: {e}")


# ── Scroll commands from remote ──
# The remote posts scroll commands; the teleprompter polls and executes them.
scroll_command = {"action": None, "version": 0}  # action: "up", "down", "top"
scroll_lock = threading.Lock()

# ── Network info (set at startup) ──
local_ip = "127.0.0.1"
server_port = DEFAULT_PORT
public_url = None  # set via --public-url for ngrok/tunnel access


def _activate_slideshow(slide_num):
    """Set the slideshow as active and jump to the given slide."""
    global slideshow_active, current_slide, total_slides
    slideshow_active = True
    current_slide = slide_num
    total_slides = len(script_sections)


def parse_script(filepath):
    """Parse a markdown script file into sections split by SLIDE markers.
    Returns (sections, titles) where titles[i] is the marker text for section i."""
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

    # Find all marker lines (for titles)
    markers = marker.findall(text)

    # Split on marker lines
    parts = marker.split(text)

    # parts[0] = text before first SLIDE marker (preamble)
    # parts[1..] = text after each SLIDE marker
    if len(parts) < 2:
        return [text.strip()], [""]

    sections = []
    titles = []
    for i, part in enumerate(parts[1:]):
        sections.append(part.strip())
        # Clean up marker text to extract title
        raw = markers[i] if i < len(markers) else ""
        # Strip markdown formatting: #, *, [, ], --, —
        clean = re.sub(r'^[#*\-\s\[]*', '', raw)
        clean = re.sub(r'[\]*]*$', '', clean)
        clean = clean.strip()
        titles.append(clean)

    return sections, titles


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
    # Trigger immediate screenshot for portrait preview
    if platform.system() == "Darwin":
        threading.Thread(target=_capture_slideshow_screenshot, daemon=True).start()


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
        # Ensure starting slide is reset to 1 (demo mode may have changed it)
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  activate\n'
            '  if (count of presentations) > 0 then\n'
            '    set thePresentation to active presentation\n'
            '    set theSettings to slide show settings of thePresentation\n'
            '    set starting slide of theSettings to 1\n'
            '    set ending slide of theSettings to count of slides of thePresentation\n'
            '    run slide show theSettings\n'
            '  end if\n'
            'end tell'
        )
    elif action == "stop":
        # End the slideshow
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  if (count of slide show windows) > 0 then\n'
            '    exit slide show slide show view of slide show window 1\n'
            '  end if\n'
            'end tell'
        )
    elif action == "check":
        # Check if slideshow is running, return "running" or "stopped"
        script = (
            'tell application "Microsoft PowerPoint"\n'
            '  if (count of slide show windows) > 0 then\n'
            '    return "running"\n'
            '  else\n'
            '    return "stopped"\n'
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


_screenshot_display = None  # cached: integer display number for -D flag


def _detect_display_count():
    """Return the number of connected displays using system_profiler."""
    try:
        result = subprocess.run(
            ["system_profiler", "SPDisplaysDataType"],
            capture_output=True, text=True, timeout=5
        )
        # Count lines containing "Resolution:" — one per display
        count = sum(1 for line in result.stdout.splitlines() if "Resolution:" in line)
        return max(count, 1)
    except Exception:
        return 1


def _try_screencapture(cmd, tmp_write, tmp_path):
    """Run a screencapture command. Returns True if a valid image was produced."""
    try:
        if os.path.exists(tmp_write):
            os.unlink(tmp_write)
        result = subprocess.run(cmd, capture_output=True, timeout=5)
        if result.returncode == 0 and os.path.exists(tmp_write):
            fsize = os.path.getsize(tmp_write)
            if fsize > 100:
                os.replace(tmp_write, tmp_path)
                return True
    except Exception:
        pass
    return False


def _set_screenshot_display(display_num):
    """Set which display to capture for the slide preview. None = auto-detect."""
    global _screenshot_display
    _screenshot_display = display_num
    label = "auto-detect" if display_num is None else f"display {display_num}"
    print(f"  Screenshot display set to: {label}")


def _capture_slideshow_screenshot():
    """Capture a screenshot of the configured display.

    Uses _screenshot_display if set (via Controls panel or auto-detect).
    Auto-detection tries display 2 first, then 3, etc., then 1 as fallback.
    """
    global live_screenshot_path, _screenshot_display
    if platform.system() != "Darwin":
        return

    import tempfile
    tmp_path = os.path.join(tempfile.gettempdir(), "teleprompter_live.jpg")
    tmp_write = tmp_path + ".tmp"

    # If we already know which display to use, use it
    if _screenshot_display is not None:
        if _try_screencapture(
            ["screencapture", "-x", "-t", "jpg", "-D", str(_screenshot_display), tmp_write],
            tmp_write, tmp_path
        ):
            with live_screenshot_lock:
                live_screenshot_path = tmp_path
        return

    # Auto-detect: try each display, preferring non-main (display 2+)
    num_displays = _detect_display_count()
    print(f"  Detecting displays: {num_displays} found")

    # Try secondary displays first (2, 3, ...), then main (1) as fallback
    display_order = list(range(2, num_displays + 1)) + [1]
    for d in display_order:
        if _try_screencapture(
            ["screencapture", "-x", "-t", "jpg", "-D", str(d), tmp_write],
            tmp_write, tmp_path
        ):
            _screenshot_display = d
            with live_screenshot_lock:
                live_screenshot_path = tmp_path
            label = "secondary" if d > 1 else "main (single display)"
            print(f"  Using display {d} ({label}) for slide preview")
            return

    print("  screencapture: all displays failed")


def _run_screenshot_loop():
    """Background loop that captures slideshow screenshots every 2 seconds."""
    global _screenshot_display
    import time as _time
    capture_count = 0
    while True:
        _time.sleep(2)
        if slideshow_active:
            _capture_slideshow_screenshot()
            capture_count += 1
            if capture_count <= 3:
                with live_screenshot_lock:
                    path = live_screenshot_path
                if path and os.path.exists(path):
                    fsize = os.path.getsize(path)
                    print(f"  Screenshot #{capture_count}: {fsize} bytes (display={_screenshot_display})")
                else:
                    print(f"  Screenshot #{capture_count}: FAILED - no file produced")
        else:
            with live_screenshot_lock:
                live_screenshot_path = None
            _screenshot_display = None  # re-detect next slideshow
            capture_count = 0


def _demo_toggle():
    """Toggle demo mode: switch between PowerPoint slideshow and Terminal.

    Enter demo mode:
      1. Open Terminal.app, cd to demo dir, type the demo command (not executed)
      2. Leave the slideshow RUNNING — just bring Terminal to front over it
      3. Audience sees Terminal fullscreen; slideshow stays alive behind it

    Exit demo mode:
      1. Stop demo process and close Terminal
      2. Activate PowerPoint — slideshow is still running on the same slide
    """
    global demo_mode, demo_slide

    if platform.system() != "Darwin":
        return False, "Demo mode switching is only supported on macOS"

    if not demo_script:
        return False, "No demo script found in the talk directory"

    demo_cmd = f"python3 {Path(demo_script).name}"
    abs_demo_dir = str(Path(demo_dir).resolve())

    if not demo_mode:
        # ── ENTER DEMO MODE ──
        demo_slide = current_slide

        # Open Terminal over the running slideshow (do NOT exit slideshow).
        # The slideshow keeps running behind Terminal — when we come back
        # we just activate PowerPoint and it's on the exact same slide.
        script = (
            # Open Terminal, cd to demo directory, clear screen
            'tell application "Terminal"\n'
            '  activate\n'
            f'  do script "cd {abs_demo_dir} && clear"\n'
            '  delay 0.3\n'
            'end tell\n'
            # Type the demo command without executing it
            'tell application "System Events"\n'
            '  tell process "Terminal"\n'
            '    set frontmost to true\n'
            f'    keystroke "{demo_cmd}"\n'
            '  end tell\n'
            'end tell\n'
            # Ensure Terminal stays frontmost
            'delay 0.2\n'
            'tell application "Terminal"\n'
            '  activate\n'
            'end tell\n'
        )

        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, timeout=15, text=True
            )
            if result.returncode != 0 and result.stderr.strip():
                return False, result.stderr.strip()
            demo_mode = True
            return True, "entered"
        except subprocess.TimeoutExpired:
            return False, "AppleScript timed out"
        except Exception as e:
            return False, str(e)

    else:
        # ── EXIT DEMO MODE ──
        # Step 1: Kill the demo process and close Terminal cleanly.
        # Ctrl-C stops running process, then 'exit' closes the shell
        # so Terminal's close doesn't trigger a confirmation dialog.
        term_script = (
            'tell application "System Events"\n'
            '  tell process "Terminal"\n'
            '    set frontmost to true\n'
            '    keystroke "c" using control down\n'
            '    delay 0.3\n'
            '    keystroke "exit"\n'
            '    key code 36\n'  # Return key
            '  end tell\n'
            'end tell\n'
            'delay 0.5\n'
        )

        try:
            subprocess.run(
                ["osascript", "-e", term_script],
                capture_output=True, timeout=10, text=True
            )
        except Exception:
            pass  # Best-effort Terminal cleanup

        # Step 2: Just activate PowerPoint. The slideshow never stopped,
        # so it's still running on the exact slide where we left it.
        resume_script = (
            'tell application "Microsoft PowerPoint"\n'
            '  activate\n'
            'end tell\n'
        )

        try:
            result = subprocess.run(
                ["osascript", "-e", resume_script],
                capture_output=True, timeout=10, text=True
            )
            demo_mode = False

            if result.returncode != 0 and result.stderr.strip():
                print(f"  Warning: resume failed: {result.stderr.strip()}")

            return True, "exited"
        except subprocess.TimeoutExpired:
            demo_mode = False
            return False, "AppleScript timed out"
        except Exception as e:
            demo_mode = False
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
    def handle(self):
        """Override to catch broken connections so they never crash the server."""
        try:
            super().handle()
        except (ConnectionResetError, BrokenPipeError, ConnectionAbortedError):
            pass  # client disconnected mid-request, harmless

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
            if public_url:
                remote_url = public_url.rstrip("/") + "/remote"
            else:
                remote_url = f"http://{local_ip}:{server_port}/remote"
            page = QR_PAGE.replace("{{REMOTE_URL}}", remote_url)
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(page.encode("utf-8"))

        elif self.path == "/api/remote-url":
            if public_url:
                remote_url = public_url.rstrip("/") + "/remote"
            else:
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

            title = slide_titles[section_idx] if section_idx < len(slide_titles) else ""

            state = {
                "slide": current_slide,
                "totalSlides": total_slides,
                "slideTitle": title,
                "scriptHtml": md_to_html(script_text),
                "totalSections": len(script_sections),
                "active": slideshow_active,
                "mode": mode,
                "demoMode": demo_mode,
                "demoAvailable": demo_script is not None,
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
                # Trigger immediate screenshot so portrait preview updates quickly
                threading.Thread(target=_capture_slideshow_screenshot, daemon=True).start()
            self._json_response({"ok": ok, "error": result if not ok else None})

        elif self.path == "/api/ppt/prev":
            ok, result = _ppt_applescript("prev")
            if ok and result and "/" in result:
                idx, tot = result.split("/", 1)
                _activate_slideshow(int(idx))
                threading.Thread(target=_capture_slideshow_screenshot, daemon=True).start()
            self._json_response({"ok": ok, "error": result if not ok else None})

        elif self.path == "/api/ppt/start":
            ok, err = _ppt_applescript("start")
            if ok:
                # Activate teleprompter on slide 1 since VBA won't fire for the initial slide
                _activate_slideshow(1)
            self._json_response({"ok": ok, "error": err})

        elif self.path == "/api/ppt/stop":
            ok, err = _ppt_applescript("stop")
            if ok:
                _receive_stopped()
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

        elif self.path == "/api/demo/toggle":
            ok, result = _demo_toggle()
            self._json_response({
                "ok": ok,
                "demoMode": demo_mode,
                "error": result if not ok else None,
                "result": result if ok else None,
            })

        elif self.path == "/api/demo/state":
            self._json_response({
                "demoMode": demo_mode,
                "demoScript": Path(demo_script).name if demo_script else None,
                "available": demo_script is not None,
            })

        elif self.path == "/api/slide-image-debug":
            # Debug endpoint to check screenshot pipeline status
            import tempfile
            tmp_path = os.path.join(tempfile.gettempdir(), "teleprompter_live.jpg")
            with live_screenshot_lock:
                current_path = live_screenshot_path
            info = {
                "slideshow_active": slideshow_active,
                "live_screenshot_path": current_path,
                "tmp_path_exists": os.path.exists(tmp_path),
                "tmp_path_size": os.path.getsize(tmp_path) if os.path.exists(tmp_path) else 0,
                "platform": platform.system(),
                "capture_display": _screenshot_display,
                "display_count": _detect_display_count(),
            }
            self._json_response(info)

        elif self.path.startswith("/api/slide-image"):
            # Serve live screenshot of the PowerPoint slideshow
            with live_screenshot_lock:
                path = live_screenshot_path
            try:
                if path:
                    with open(path, "rb") as f:
                        img_data = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "image/jpeg")
                    self.send_header("Content-Length", str(len(img_data)))
                    self.send_header("Cache-Control", "no-cache, no-store")
                    self.end_headers()
                    self.wfile.write(img_data)
                else:
                    self.send_response(404)
                    self.send_header("Content-Type", "text/plain")
                    self.end_headers()
                    msg = f"No screenshot available. active={slideshow_active}, path={path}"
                    self.wfile.write(msg.encode())
            except (ConnectionResetError, BrokenPipeError):
                pass  # browser disconnected mid-transfer, harmless
            except (FileNotFoundError, OSError):
                try:
                    self.send_response(404)
                    self.send_header("Content-Type", "text/plain")
                    self.end_headers()
                    self.wfile.write(b"Screenshot file was removed during read")
                except (ConnectionResetError, BrokenPipeError):
                    pass

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
                                "autoScroll", "scrollSpeed", "mirror",
                                "highlightLevel", "highlightLines",
                                "panelHeightPct", "uiScaleLevel"):
                        if key in updates:
                            display_settings[key] = updates[key]
                    display_settings["settingsVersion"] += 1
                    result = dict(display_settings)

                # Persist to disk in background (non-blocking)
                threading.Thread(target=_save_settings, daemon=True).start()

                self._json_response(result)
            except Exception:
                self.send_response(400)
                self.end_headers()

        elif self.path == "/api/screenshot-display":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(length)
                data = json.loads(body)
                d = int(data.get("display", 0))
                if d == 0:
                    # Reset to auto-detect
                    _set_screenshot_display(None)
                else:
                    _set_screenshot_display(d)
                # Persist the display choice
                with settings_lock:
                    display_settings["screenshotDisplay"] = d
                threading.Thread(target=_save_settings, daemon=True).start()
                # Trigger an immediate capture with the new display
                threading.Thread(target=_capture_slideshow_screenshot, daemon=True).start()
                self._json_response({"ok": True, "display": d})
            except Exception as e:
                self._json_response({"ok": False, "error": str(e)}, 400)

        elif self.path == "/api/focus-browser":
            # Use AppleScript to find the browser window with "Teleprompter"
            # in its title and raise that specific window.
            if platform.system() == "Darwin":
                # First, figure out which browsers are actually running
                running = []
                for name in ("Google Chrome", "Safari", "Firefox"):
                    try:
                        r = subprocess.run(
                            ["osascript", "-e",
                             f'tell application "System Events" to '
                             f'return exists process "{name}"'],
                            capture_output=True, text=True, timeout=3
                        )
                        if r.returncode == 0 and "true" in r.stdout.lower():
                            running.append(name)
                    except Exception:
                        pass

                focused = False
                err_detail = ""
                for browser in running:
                    try:
                        if browser == "Google Chrome":
                            script = '''
tell application "Google Chrome"
    set winCount to count of windows
    repeat with i from 1 to winCount
        set w to window i
        set tabCount to count of tabs of w
        repeat with j from 1 to tabCount
            if title of tab j of w contains "Teleprompter" then
                set active tab index of w to j
                set index of w to 1
                activate
                return "found"
            end if
        end repeat
    end repeat
    -- Tab not found, just activate Chrome anyway
    activate
    return "activated"
end tell
'''
                        elif browser == "Safari":
                            script = '''
tell application "Safari"
    set winCount to count of windows
    repeat with i from 1 to winCount
        set w to window i
        set tabCount to count of tabs of w
        repeat with j from 1 to tabCount
            if name of tab j of w contains "Teleprompter" then
                set current tab of w to tab j of w
                set index of w to 1
                activate
                return "found"
            end if
        end repeat
    end repeat
    activate
    return "activated"
end tell
'''
                        else:  # Firefox — no tab-level AppleScript API
                            script = '''
tell application "Firefox" to activate
return "activated"
'''
                        result = subprocess.run(
                            ["osascript", "-e", script],
                            capture_output=True, text=True, timeout=5
                        )
                        if result.returncode == 0:
                            status = result.stdout.strip()
                            # AppleScript activate can be flaky — also use
                            # 'open -a' which is more reliable at bringing
                            # the app to the foreground on macOS.
                            subprocess.run(
                                ["open", "-a", browser],
                                capture_output=True, timeout=3
                            )
                            self._json_response({
                                "ok": True,
                                "browser": browser,
                                "matched": status == "found"
                            })
                            focused = True
                            break
                        else:
                            err_detail = result.stderr.strip()
                    except Exception as e:
                        err_detail = str(e)

                if not focused:
                    # Last resort: try open -a for each running browser
                    for browser in running:
                        try:
                            subprocess.run(
                                ["open", "-a", browser],
                                capture_output=True, timeout=3
                            )
                            self._json_response({
                                "ok": True,
                                "browser": browser,
                                "matched": False
                            })
                            focused = True
                            break
                        except Exception:
                            pass
                if not focused:
                    msg = "No running browser found" if not running else f"AppleScript error: {err_detail}"
                    self._json_response({"ok": False, "error": msg})
            else:
                self._json_response({"ok": False, "error": "macOS only"})

        else:
            self.send_response(404)
            self.end_headers()


def md_to_html(text):
    """Minimal markdown to HTML conversion."""
    import html as html_mod
    text = html_mod.escape(text)
    # Restore allowed HTML tags for Q&A expand/collapse sections
    for tag in ("details", "summary", "div", "span", "hr"):
        text = text.replace(f"&lt;{tag}&gt;", f"<{tag}>")
        text = text.replace(f"&lt;/{tag}&gt;", f"</{tag}>")
        # Handle tags with attributes (e.g. <div class="qa-index">)
        # After html.escape(), quotes become &quot; so we match broadly
        text = re.sub(
            rf"&lt;({tag}\s+.*?)&gt;",
            lambda m: "<" + m.group(1).replace("&quot;", '"') + ">",
            text,
        )
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

  /* Keep-alive pixel — continuous animation prevents screen sleep */
  .keep-alive {
    position: fixed;
    bottom: 0;
    right: 0;
    width: 1px;
    height: 1px;
    opacity: 0.01;
    animation: keepAlive 1s infinite alternate;
    pointer-events: none;
  }
  @keyframes keepAlive {
    from { transform: translateX(0); }
    to { transform: translateX(1px); }
  }
</style>
</head>
<body>
<div class="keep-alive"></div>

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

<!-- Focus browser (activates the teleprompter browser window via AppleScript) -->
<button class="big-btn" id="focusBrowserBtn" style="background:#555;border-color:#666;color:#fff;" ontouchend="focusBrowser(event)" onclick="focusBrowser(event)">
  &#x1F5A5; Focus Teleprompter Screen
</button>

<!-- Demo mode toggle -->
<button class="big-btn" id="demoBtn" style="background:#7c3aed;border-color:#7c3aed;color:#fff;" ontouchend="toggleDemo(event)" onclick="toggleDemo(event)">
  &#x1F4BB; Switch to Demo
</button>

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

<!-- Highlight bar -->
<div class="section-label">Highlight Bar <span id="hlRemoteVal" style="float:right;opacity:0.7;">Off</span></div>
<div class="ctrl-row">
  <input type="range" min="0" max="100" value="0" id="highlightRemoteSlider"
    style="width:100%;accent-color:#f0c040;height:36px;" />
</div>
<div class="ctrl-row" style="margin-top:4px;">
  <span style="flex:1;font-size:0.95rem;color:#aaa;">HL Lines</span>
  <button class="ctrl-btn" ontouchend="adjHlLines(-1,event)" onclick="adjHlLines(-1,event)">-</button>
  <span id="hlLinesRemoteVal" style="min-width:28px;text-align:center;font-size:1.1rem;">3</span>
  <button class="ctrl-btn" ontouchend="adjHlLines(1,event)" onclick="adjHlLines(1,event)">+</button>
</div>

<div class="status-bar" id="statusBar">Connecting...</div>

<script>
var settings = {};
var demoActive = false;

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

function focusBrowser(e) {
  stopEvent(e);
  document.getElementById('statusBar').textContent = 'Focusing...';
  fetch('/api/focus-browser', { method: 'POST' })
    .then(function(r) { return r.json(); })
    .then(function(data) {
      console.log('focus response', JSON.stringify(data));
      if (data.ok) {
        var msg = data.matched ? 'Focused teleprompter in ' + data.browser
                               : 'Activated ' + data.browser + ' (tab not matched)';
        document.getElementById('statusBar').textContent = msg;
      } else {
        document.getElementById('statusBar').textContent = 'Focus failed: ' + (data.error || 'unknown');
      }
    })
    .catch(function(err) {
      console.log('focus error', err);
      document.getElementById('statusBar').textContent = 'Focus request failed';
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

function toggleDemo(e) {
  stopEvent(e);
  var btn = document.getElementById('demoBtn');
  btn.style.opacity = '0.5';
  btn.textContent = 'Switching...';
  fetch('/api/demo/toggle')
    .then(function(r) { return r.json(); })
    .then(function(data) {
      btn.style.opacity = '1';
      if (!data.ok && data.error) {
        document.getElementById('statusBar').textContent = 'Error: ' + data.error;
        btn.innerHTML = '&#x1F4BB; Switch to Demo';
      } else {
        updateDemoBtn(data.demoMode);
      }
    })
    .catch(function() {
      btn.style.opacity = '1';
      btn.innerHTML = '&#x1F4BB; Switch to Demo';
      document.getElementById('statusBar').textContent = 'Connection error';
    });
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

function updateDemoBtn(isDemo) {
  var btn = document.getElementById('demoBtn');
  if (isDemo) {
    btn.style.background = '#dc2626';
    btn.style.borderColor = '#dc2626';
    btn.innerHTML = '&#x25B6; Back to Slides';
  } else {
    btn.style.background = '#7c3aed';
    btn.style.borderColor = '#7c3aed';
    btn.innerHTML = '&#x1F4BB; Switch to Demo';
  }
}

function updateUI() {
  document.getElementById('fontVal').textContent = (settings.fontSize || 2.8).toFixed(1);
  document.getElementById('widthVal').textContent = (settings.textWidth || 1800);
  document.getElementById('wordSpVal').textContent = (settings.wordSpacing || 0) + 'px';
  var hlSlider = document.getElementById('highlightRemoteSlider');
  var hlVal = document.getElementById('hlRemoteVal');
  var hl = settings.highlightLevel || 0;
  if (hlSlider && parseInt(hlSlider.value) !== hl) hlSlider.value = hl;
  if (hlVal) hlVal.textContent = hl === 0 ? 'Off' : hl + '%';
  var hlLinesVal = document.getElementById('hlLinesRemoteVal');
  if (hlLinesVal) hlLinesVal.textContent = settings.highlightLines || 3;
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
      // Sync demo button state
      if (typeof data.demoMode !== 'undefined') {
        if (data.demoMode !== demoActive) {
          demoActive = data.demoMode;
          updateDemoBtn(demoActive);
        }
        // Hide demo button if no demo script available
        var demoBtn = document.getElementById('demoBtn');
        if (demoBtn) demoBtn.style.display = data.demoAvailable ? 'flex' : 'none';
      }
      document.getElementById('statusBar').textContent = 'Connected';
    })
    .catch(function() {
      document.getElementById('statusBar').textContent = 'Connection lost - retrying...';
    });
}

setInterval(poll, 500);
poll();

// ── Highlight slider on remote ──
document.getElementById('highlightRemoteSlider').addEventListener('input', function() {
  postSettings({ highlightLevel: parseInt(this.value) });
});

function adjHlLines(delta, e) {
  stopEvent(e);
  var cur = settings.highlightLines || 3;
  var next = Math.max(1, Math.min(8, cur + delta));
  settings.highlightLines = next;
  document.getElementById('hlLinesRemoteVal').textContent = next;
  postSettings({ highlightLines: next });
}

// ── Keep screen awake ──
// Method 1: Wake Lock API (modern browsers)
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
document.addEventListener('visibilitychange', function() {
  if (document.visibilityState === 'visible') requestWakeLock();
});

// Method 2: Hidden canvas redraws every second to signal screen activity
(function() {
  var c = document.createElement('canvas');
  c.width = 2; c.height = 2;
  c.style.cssText = 'position:fixed;bottom:0;right:0;width:1px;height:1px;opacity:0.01;pointer-events:none;';
  document.body.appendChild(c);
  var ctx = c.getContext('2d');
  setInterval(function() {
    ctx.fillStyle = 'rgba(' + (Math.random()*255|0) + ',0,0,0.01)';
    ctx.fillRect(0, 0, 2, 2);
  }, 1000);
})();

// Method 3: Periodic fetch + title toggle to keep browser active
setInterval(function() {
  document.title = document.title === 'Teleprompter Remote' ?
    'Teleprompter Remote ' : 'Teleprompter Remote';
}, 10000);
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
    position: relative;
    touch-action: none;  /* let JS handle all touch/pointer gestures */
  }
  .teleprompter .inner {
    max-width: 1800px;
    margin: 0 auto;
    transition: opacity 0.12s;
    padding-bottom: 100vh;  /* allows scrolling text fully off the top */
  }
  .teleprompter p { margin-bottom: 1.1em; }
  .teleprompter strong { color: var(--accent); }
  .teleprompter em { color: #f0c040; }
  .teleprompter.mirror { transform: scaleX(-1); }

  /* Highlight bar overlaying the first lines of text in the teleprompter.
     Positioned dynamically by JS to align with the teleprompter scroll area. */
  .highlight-bar {
    display: none;
    position: sticky;
    top: 0;
    left: 0;
    right: 0;
    height: 5.7em;       /* default 3 lines × 1.9 line-height; overridden by JS */
    margin-bottom: -5.7em;
    background: rgba(255, 245, 140, var(--hl-opacity, 0));
    border-bottom: 2px solid rgba(255, 230, 50, calc(var(--hl-opacity, 0) * 2.5));
    pointer-events: none;
    z-index: 50;
  }
  body.highlight-on .highlight-bar {
    display: block;
  }

  /* Scroll remaining indicator */
  .scroll-indicator {
    position: fixed;
    left: 12px;
    top: 50%;
    transform: translateY(-50%);
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 4px;
    z-index: 20;
    pointer-events: none;
    opacity: 0;
    transition: opacity 0.3s;
  }
  .scroll-indicator.visible {
    opacity: 1;
  }
  .scroll-indicator .scroll-pct {
    font-size: 1.8rem;
    color: #ccc;
    font-weight: 700;
    background: rgba(30, 30, 30, 0.75);
    padding: 8px 14px;
    border-radius: 8px;
    white-space: nowrap;
  }
  .scroll-indicator .scroll-arrows {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 0;
    line-height: 1;
  }
  .scroll-indicator .scroll-arrows .arr {
    font-size: 2.2rem;
    color: #ccc;
    line-height: 0.7;
  }

  /* Highlight slider styling in control panel */
  .cp-highlight-slider {
    width: 100%;
    margin: 2px 0;
    accent-color: #f0c040;
    cursor: pointer;
  }

  /* Q&A expand/collapse index */
  .qa-index { margin-top: 0.5em; }
  .qa-index h2 { font-size: 1.4em; color: var(--accent); margin-bottom: 0.4em; }
  .qa-index hr { border: none; border-top: 1px solid var(--border); margin: 0.5em 0; }
  .qa-index .qa-section-title { font-size: 1.1em; color: #2E75B6; margin: 0.6em 0 0.2em; }
  .qa-index details {
    margin: 0.15em 0;
    border-left: 3px solid var(--border);
    padding-left: 0.5em;
    transition: border-color 0.2s;
  }
  .qa-index details[open] { border-left-color: var(--accent); }
  .qa-index summary {
    cursor: pointer;
    font-size: 0.75em;
    font-weight: 600;
    color: var(--text);
    padding: 0.25em 0;
    list-style: none;
    display: flex;
    align-items: baseline;
    gap: 0.4em;
  }
  .qa-index summary::before {
    content: '\25B6';
    font-size: 0.6em;
    color: var(--dim);
    transition: transform 0.2s;
    flex-shrink: 0;
  }
  .qa-index details[open] > summary::before { transform: rotate(90deg); color: var(--accent); }
  .qa-index summary::-webkit-details-marker { display: none; }
  .qa-index .qa-answer {
    font-size: 0.85em;
    font-weight: 400;
    color: #ccc;
    line-height: 1.6;
    padding: 0.3em 0 0.5em 0;
  }

  /* ── Portrait mode: slide info panel at bottom ── */
  .slide-panel {
    display: none;
    flex-shrink: 0;
    background: #0d0d14;
    border-top: 2px solid var(--border);
    padding: 0 0 12px;
    position: relative;
  }
  .slide-panel-drag {
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 8px 0 4px;
    cursor: ns-resize;
    touch-action: none;
    user-select: none;
    -webkit-user-select: none;
  }
  .slide-panel-drag .drag-pill {
    width: 48px;
    height: 5px;
    border-radius: 3px;
    background: #555;
    transition: background 0.15s;
  }
  .slide-panel-drag:hover .drag-pill,
  .slide-panel-drag:active .drag-pill { background: #888; }
  .slide-panel-preview {
    text-align: center;
    margin-bottom: 10px;
    padding: 0 4px;
  }
  .slide-panel-preview img {
    width: 100%;
    max-height: var(--panel-img-max-h, 30vh);
    border-radius: 4px;
    border: 1px solid var(--border);
    background: #000;
    object-fit: contain;
  }
  .slide-panel-preview img[src=""] { display: none; }
  .slide-panel-preview img:not([src=""]):not([src]) { display: none; }
  .slide-panel-info {
    text-align: center;
    margin-bottom: 10px;
    padding: 0 16px;
  }
  .slide-panel .slide-panel-num {
    font-size: 1.6rem;
    font-weight: 800;
    color: var(--accent);
    margin-bottom: 2px;
  }
  .slide-panel .slide-panel-title {
    font-size: 1.1rem;
    color: var(--dim);
    font-weight: 500;
  }
  .slide-panel .slide-panel-nav {
    display: flex;
    gap: 12px;
    justify-content: center;
    padding: 0 16px;
  }
  .slide-panel .slide-panel-btn {
    padding: 10px 28px;
    border-radius: 10px;
    border: 2px solid #1a6b3a;
    background: #1a6b3a;
    color: #fff;
    font-size: 1rem;
    font-weight: 700;
    cursor: pointer;
    touch-action: manipulation;
    transition: 0.12s;
  }
  .slide-panel .slide-panel-btn:hover { background: #238c4e; border-color: #238c4e; }
  .slide-panel .slide-panel-btn:active { transform: scale(0.95); }

  body.portrait .slide-panel { display: block; }
  body.portrait .teleprompter { flex: 1; }
  body.portrait .footer { display: none; }

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
    max-height: calc(100vh - 80px);
    display: flex;
    flex-direction: column;
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
    flex-shrink: 0;
  }
  .control-panel.collapsed .cp-header { border-bottom: none; }
  .cp-header:hover { color: var(--text); }
  .cp-toggle { font-size: 0.85rem; color: var(--dim); }

  .cp-body { padding: 16px 18px; overflow-y: auto; flex: 1; min-height: 0; }

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

  .demo-toggle {
    width: 100%;
    padding: 14px;
    margin-bottom: 14px;
    border-radius: 10px;
    border: 2px solid #7c3aed;
    background: #7c3aed;
    color: #fff;
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
  .demo-toggle:hover { background: #6d28d9; border-color: #6d28d9; }
  .demo-toggle.active { background: #dc2626; border-color: #dc2626; }
  .demo-toggle.active:hover { background: #b91c1c; border-color: #b91c1c; }

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
  <div class="scroll-indicator" id="scrollIndicator">
    <div class="scroll-pct" id="scrollPct"></div>
    <div class="scroll-arrows" id="scrollArrows"></div>
  </div>
  <div class="teleprompter" id="teleprompter">
    <div class="highlight-bar" id="highlightBar"></div>
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

      <!-- Portrait mode -->
      <div class="cp-row">
        <span class="cp-label">Portrait</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="portraitBtn">Off</button>
        </div>
      </div>

      <!-- Capture display -->
      <div class="cp-row">
        <span class="cp-label">Capture</span>
        <div class="cp-btn-group">
          <button class="cp-btn" id="dispPrev" style="font-size:1.1rem;">&#x25C0;</button>
          <span class="cp-value" id="dispValue">Auto</span>
          <button class="cp-btn" id="dispNext" style="font-size:1.1rem;">&#x25B6;</button>
        </div>
      </div>

      <!-- Highlight bar -->
      <div class="cp-row" style="flex-wrap:wrap;">
        <span class="cp-label">Highlight</span>
        <span class="cp-value" id="highlightValue" style="margin-left:auto;font-size:0.85rem;">Off</span>
      </div>
      <div style="padding:0 18px 8px;">
        <input type="range" min="0" max="100" value="0" class="cp-highlight-slider" id="highlightSlider" />
      </div>
      <div class="cp-row">
        <span class="cp-label">HL Lines</span>
        <div class="cp-btn-group">
          <button class="cp-btn" onclick="changeHighlightLines(-1)">&#9664;</button>
          <span class="cp-value" id="hlLinesValue">3</span>
          <button class="cp-btn" onclick="changeHighlightLines(1)">&#9654;</button>
        </div>
      </div>

      <!-- UI Scale -->
      <div class="cp-row">
        <span class="cp-label">UI Scale</span>
        <div class="cp-btn-group">
          <button class="cp-btn" onclick="changeUiScale(-1)">&#x2212;</button>
          <span class="cp-value" id="uiScaleValue">1x</span>
          <button class="cp-btn" onclick="changeUiScale(1)">+</button>
        </div>
      </div>

      <!-- Demo mode toggle -->
      <button class="demo-toggle" id="demoToggleBtn">
        &#x1F4BB; Switch to Demo
      </button>

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

<!-- Slide info panel (portrait mode) -->
<div class="slide-panel" id="slidePanel">
  <div class="slide-panel-drag" id="slidePanelDrag">
    <div class="drag-pill"></div>
  </div>
  <div class="slide-panel-preview" id="slidePanelPreview">
    <img id="slidePanelImg" alt="" />
  </div>
  <div class="slide-panel-info">
    <div class="slide-panel-num" id="slidePanelNum">&mdash;</div>
    <div class="slide-panel-title" id="slidePanelTitle"></div>
  </div>
  <div class="slide-panel-nav">
    <button class="slide-panel-btn" id="panelPrevSlide">&#x25C0; Prev</button>
    <button class="slide-panel-btn" id="panelNextSlide">Next &#x25B6;</button>
  </div>
</div>

<div class="footer" id="footer">
  <span><kbd>D</kbd> Demo mode</span>
  <span><kbd>P</kbd> Portrait</span>
  <span><kbd>H</kbd> Highlight</span>
  <span><kbd>T</kbd> Timer</span>
  <span>Open <strong>/remote</strong> on your phone to control without touching this screen</span>
  <span id="modeHint">Syncs via VBA</span>
</div>

<script>
var lastSlide = -1;
var slideshow_active_flag = false;
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

// ── Scroll remaining indicator ──
function updateScrollIndicator() {
  var tp = document.getElementById('teleprompter');
  var indicator = document.getElementById('scrollIndicator');
  var pctEl = document.getElementById('scrollPct');
  var arrowsEl = document.getElementById('scrollArrows');

  // Measure real content height excluding the 100vh padding-bottom on .inner.
  // Use .inner's scrollHeight minus its computed padding-bottom.
  var innerEl = document.getElementById('scriptContent');
  var innerPadBottom = parseFloat(getComputedStyle(innerEl).paddingBottom) || 0;
  // realContentBottom: position of the bottom of real text relative to the
  // teleprompter scroll container (includes teleprompter's padding-top).
  var realContentBottom = innerEl.offsetTop + innerEl.scrollHeight - innerPadBottom;

  var viewportHeight = tp.clientHeight;
  var scrollPos = tp.scrollTop;

  if (realContentBottom <= viewportHeight) {
    // All text fits on screen — no indicator needed
    indicator.classList.remove('visible');
  } else {
    // How far through the real content have we scrolled?
    var maxScroll = realContentBottom - viewportHeight;
    var pct = Math.round((scrollPos / maxScroll) * 100);
    pct = Math.max(0, Math.min(100, pct));
    indicator.classList.add('visible');
    pctEl.textContent = pct + '%';

    // Number of arrows = number of remaining screens of text (rounded up)
    var hiddenBelow = realContentBottom - (scrollPos + viewportHeight);
    var screensLeft = Math.ceil(hiddenBelow / viewportHeight);
    screensLeft = Math.max(0, Math.min(screensLeft, 10));  // cap at 10

    // Build arrow HTML (one ▼ per remaining screen)
    var arrowHtml = '';
    for (var a = 0; a < screensLeft; a++) {
      arrowHtml += '<div class="arr">\u25BC</div>';
    }
    arrowsEl.innerHTML = arrowHtml;
  }
}

document.getElementById('teleprompter').addEventListener('scroll', updateScrollIndicator);
// Also update on window resize and after content changes
window.addEventListener('resize', updateScrollIndicator);

// ── Text width ──
var textWidth = 1800;
var WIDTH_STEP = 200, WIDTH_MIN = 600, WIDTH_MAX = 3000;

// ── Word spacing ──
var wordSpacing = 0;  // px
var WORDSP_STEP = 2, WORDSP_MIN = -4, WORDSP_MAX = 24;

// ── Demo mode ──
var demoModeActive = false;
var demoAvailable = false;

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
      mirror: mirrorOn,
      highlightLevel: highlightLevel,
      highlightLines: highlightLines,
      uiScaleLevel: uiScaleLevel
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
  if (s.mirror !== mirrorOn) {
    mirrorOn = s.mirror;
    document.getElementById('teleprompter').classList.toggle('mirror', mirrorOn);
    document.getElementById('mirrorBtn').textContent = mirrorOn ? 'On' : 'Off';
  }
  if (typeof s.highlightLevel !== 'undefined' && s.highlightLevel !== highlightLevel) {
    setHighlightLevel(s.highlightLevel);
  }
  if (typeof s.highlightLines !== 'undefined' && s.highlightLines !== highlightLines) {
    highlightLines = s.highlightLines;
    applyHighlightLines();
  }
  if (typeof s.uiScaleLevel !== 'undefined' && s.uiScaleLevel !== uiScaleLevel) {
    uiScaleLevel = s.uiScaleLevel;
    applyUiScale();
  }
  if (typeof s.screenshotDisplay !== 'undefined' && s.screenshotDisplay !== captureDisplay) {
    captureDisplay = s.screenshotDisplay;
    document.getElementById('dispValue').textContent = captureDisplay === 0 ? 'Auto' : 'Display ' + captureDisplay;
  }
  if (typeof s.panelHeightPct !== 'undefined' && s.panelHeightPct > 0) {
    var panel = document.getElementById('slidePanel');
    var h = Math.round(window.innerHeight * s.panelHeightPct / 100);
    panel.style.height = h + 'px';
    var imgMaxH = Math.max(40, h - 100);
    document.documentElement.style.setProperty('--panel-img-max-h', imgMaxH + 'px');
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

function toggleDemoMode() {
  var btn = document.getElementById('demoToggleBtn');
  btn.style.opacity = '0.5';
  btn.textContent = 'Switching...';
  fetch('/api/demo/toggle')
    .then(function(r) { return r.json(); })
    .then(function(data) {
      btn.style.opacity = '1';
      if (data.ok) {
        demoModeActive = data.demoMode;
        updateDemoToggle();
        showTouchHint(data.demoMode ? 'Switched to demo' : 'Back to slides');
      } else {
        updateDemoToggle();
        showTouchHint('Error: ' + (data.error || 'unknown'));
      }
    })
    .catch(function() {
      btn.style.opacity = '1';
      updateDemoToggle();
      showTouchHint('Connection error');
    });
}

function updateDemoToggle() {
  var btn = document.getElementById('demoToggleBtn');
  if (!btn) return;
  if (!demoAvailable) {
    btn.style.display = 'none';
    return;
  }
  btn.style.display = 'flex';
  if (demoModeActive) {
    btn.className = 'demo-toggle active';
    btn.innerHTML = '&#x25B6; Back to Slides';
  } else {
    btn.className = 'demo-toggle';
    btn.innerHTML = '&#x1F4BB; Switch to Demo';
  }
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

// ── Drag-to-scroll with momentum ──
// Uses ONLY pointer events (unified API for mouse, touch, pen).
// CSS touch-action:none on .teleprompter prevents browser from
// handling gestures natively, so our JS gets all events reliably.
(function() {
  var tp = document.getElementById('teleprompter');
  var dragging = false;
  var pointerId = -1;
  var lastY = 0;
  var lastTime = 0;
  var velocity = 0;
  var momentumId = 0;
  var FRICTION = 0.95;
  var MIN_VEL = 0.5;

  function isControl(el) {
    if (!el) return false;
    return el.closest('.control-panel') || el.closest('.slide-panel-drag')
        || el.closest('.slide-panel-nav') || el.closest('.slide-panel-btn')
        || el.tagName === 'INPUT' || el.tagName === 'BUTTON';
  }

  function stopMomentum() {
    if (momentumId) { cancelAnimationFrame(momentumId); momentumId = 0; }
    velocity = 0;
  }

  function momentumStep() {
    velocity *= FRICTION;
    if (Math.abs(velocity) < MIN_VEL) { momentumId = 0; return; }
    tp.scrollTop += velocity;
    momentumId = requestAnimationFrame(momentumStep);
  }

  tp.addEventListener('pointerdown', function(e) {
    if (isControl(e.target)) return;
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    stopMomentum();
    dragging = true;
    pointerId = e.pointerId;
    lastY = e.clientY;
    lastTime = Date.now();
    velocity = 0;
    tp.style.cursor = 'grabbing';
    document.body.style.userSelect = 'none';
    document.body.style.webkitUserSelect = 'none';
    // Do NOT use setPointerCapture — it causes cursor drift on some drivers.
    // Instead we listen on document for move/up.
  });

  document.addEventListener('pointermove', function(e) {
    if (!dragging || e.pointerId !== pointerId) return;
    e.preventDefault();
    var y = e.clientY;
    var now = Date.now();
    var moveDy = lastY - y;
    var dt = Math.max(1, now - lastTime);
    velocity = 0.7 * (moveDy / dt * 16) + 0.3 * velocity;
    tp.scrollTop += moveDy;
    lastY = y;
    lastTime = now;
  });

  function endDrag(e) {
    if (!dragging) return;
    if (e && e.pointerId !== undefined && e.pointerId !== pointerId) return;
    if (Math.abs(velocity) > MIN_VEL) {
      momentumId = requestAnimationFrame(momentumStep);
    }
    dragging = false;
    pointerId = -1;
    tp.style.cursor = 'grab';
    document.body.style.userSelect = '';
    document.body.style.webkitUserSelect = '';
  }

  document.addEventListener('pointerup', endDrag);
  document.addEventListener('pointercancel', endDrag);

  tp.style.cursor = 'grab';
})();

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

    slideshow_active_flag = data.active;
    if (data.active) {
      dot.className = 'status-dot active';
      if (data.mode !== 'keyboard') statusText.textContent = 'Slideshow active';

      document.getElementById('slideInfo').textContent =
        'Slide ' + data.slide + ' / ' + data.totalSlides;
      updateSlidePanel(data);

      var pct = data.totalSlides > 1
        ? ((data.slide - 1) / (data.totalSlides - 1)) * 100 : 100;
      document.getElementById('progressFill').style.width = pct + '%';

      // Sync demo mode state
      if (typeof data.demoMode !== 'undefined') {
        demoModeActive = data.demoMode;
        demoAvailable = data.demoAvailable || false;
        updateDemoToggle();
      }

      if (data.slide !== lastSlide) {
        content.innerHTML = data.scriptHtml ||
          '<p style="color:var(--dim)">No script for this slide.</p>';
        tp.scrollTo(0, 0);
        lastSlide = data.slide;
        if (!timerRunning && !timerElapsed) toggleTimer();
        setTimeout(updateScrollIndicator, 50);
      }
      wasActive = true;
    } else {
      dot.className = 'status-dot waiting';
      statusText.textContent = wasActive ? 'Slideshow ended' : 'Waiting for slideshow...';
      document.getElementById('slideInfo').textContent = '\u2014';
      document.getElementById('progressFill').style.width = '0%';
      updateSlidePanel(data);
      if (lastSlide !== -1) {
        content.innerHTML =
          '<div class="waiting-msg"><div class="icon">&#x1F4E1;</div>' +
          '<p>' + (wasActive ? 'Slideshow ended.' : 'Waiting for slideshow...<br>Start your PowerPoint slideshow and the script will appear here.') + '</p></div>';
        lastSlide = -1;
        document.getElementById('scrollIndicator').classList.remove('visible');
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

// ── Portrait mode ──
var portraitOn = false;
function togglePortrait() {
  portraitOn = !portraitOn;
  document.body.classList.toggle('portrait', portraitOn);
  document.getElementById('portraitBtn').textContent = portraitOn ? 'On' : 'Off';
  if (portraitOn && slideshow_active_flag) {
    startSlideImageRefresh();
  } else {
    stopSlideImageRefresh();
  }
}

// ── Capture display picker ──
var captureDisplay = 0;  // 0 = auto
var maxDisplays = 3;     // updated from debug endpoint

function setCaptureDisplay(d) {
  captureDisplay = d;
  document.getElementById('dispValue').textContent = d === 0 ? 'Auto' : 'Display ' + d;
  fetch('/api/screenshot-display', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ display: d })
  }).catch(function(){});
}

function changeDisplay(delta) {
  var next = captureDisplay + delta;
  if (next < 0) next = maxDisplays;
  if (next > maxDisplays) next = 0;
  setCaptureDisplay(next);
}

// Fetch display count once
fetch('/api/slide-image-debug').then(function(r) { return r.json(); }).then(function(data) {
  if (data.display_count) maxDisplays = data.display_count;
}).catch(function(){});

// ── Highlight bar ──
var highlightOn = false;
var highlightLevel = 0;  // 0-100
var highlightLines = 3;  // number of text lines the bar covers (1-8)

// Highlight bar uses position:sticky inside the teleprompter — no JS positioning needed.

function applyHighlightLines() {
  // Each line of text is roughly 1.4em tall (font-size × line-height ~1.9,
  // but the bar covers from the top of the first line to the bottom of the Nth).
  // Use 1.9em per line to match the teleprompter's line-height.
  var h = (highlightLines * 1.9) + 'em';
  var bar = document.getElementById('highlightBar');
  if (bar) {
    bar.style.height = h;
    bar.style.marginBottom = '-' + h;
  }
  var label = document.getElementById('hlLinesValue');
  if (label) label.textContent = highlightLines;
}

function changeHighlightLines(delta) {
  highlightLines = Math.max(1, Math.min(8, highlightLines + delta));
  applyHighlightLines();
  pushSettings();
}

applyHighlightLines();

function setHighlightLevel(val) {
  highlightLevel = Math.max(0, Math.min(100, Math.round(val)));
  highlightOn = highlightLevel > 0;
  document.body.classList.toggle('highlight-on', highlightOn);
  // Map 0-100 to opacity 0.0-0.45 for the yellow tint
  var opacity = (highlightLevel / 100) * 0.45;
  document.documentElement.style.setProperty('--hl-opacity', opacity.toFixed(3));
  var slider = document.getElementById('highlightSlider');
  if (slider) slider.value = highlightLevel;
  var label = document.getElementById('highlightValue');
  if (label) label.textContent = highlightLevel === 0 ? 'Off' : highlightLevel + '%';
  pushSettings();
}

function toggleHighlight() {
  // Keyboard shortcut: toggle between 0 and 50%
  setHighlightLevel(highlightOn ? 0 : 50);
}

// ── UI Scale (makes control panel buttons bigger/smaller) ──
var uiScaleLevel = 0;  // -2 to +4, 0 = normal
var UI_SCALES = [0.8, 0.9, 1.0, 1.1, 1.25, 1.4, 1.6];
var UI_SCALE_LABELS = ['0.8x', '0.9x', '1x', '1.1x', '1.25x', '1.4x', '1.6x'];

function applyUiScale() {
  var idx = uiScaleLevel + 2;  // offset so level 0 maps to index 2 (1.0)
  idx = Math.max(0, Math.min(UI_SCALES.length - 1, idx));
  var scale = UI_SCALES[idx];
  var panel = document.getElementById('controlPanel');
  var slidePanel = document.getElementById('slidePanel');
  if (panel) {
    panel.style.transformOrigin = 'bottom right';
    panel.style.transform = scale === 1 ? '' : 'scale(' + scale + ')';
  }
  if (slidePanel) {
    slidePanel.style.transformOrigin = 'bottom center';
    slidePanel.style.transform = scale === 1 ? '' : 'scale(' + scale + ')';
  }
  var label = document.getElementById('uiScaleValue');
  if (label) label.textContent = UI_SCALE_LABELS[idx];
}

function changeUiScale(delta) {
  uiScaleLevel = Math.max(-2, Math.min(4, uiScaleLevel + delta));
  applyUiScale();
  pushSettings();
}

applyUiScale();

var slideImgTimer = null;
function updateSlidePanel(data) {
  var numEl = document.getElementById('slidePanelNum');
  var titleEl = document.getElementById('slidePanelTitle');
  var imgEl = document.getElementById('slidePanelImg');
  if (data.active) {
    numEl.textContent = 'Slide ' + data.slide + ' / ' + data.totalSlides;
    titleEl.textContent = data.slideTitle || '';
    // Start refreshing the live screenshot if not already running
    if (!slideImgTimer && portraitOn) startSlideImageRefresh();
  } else {
    numEl.innerHTML = '&mdash;';
    titleEl.textContent = '';
    imgEl.src = '';
    imgEl.style.display = 'none';
    stopSlideImageRefresh();
  }
}

function refreshSlideImage() {
  var imgEl = document.getElementById('slidePanelImg');
  var newImg = new Image();
  newImg.onload = function() {
    imgEl.src = newImg.src;
    imgEl.style.display = 'block';
  };
  newImg.onerror = function() {};
  // Cache-bust with timestamp
  newImg.src = '/api/slide-image?t=' + Date.now();
}

function startSlideImageRefresh() {
  if (slideImgTimer) return;
  refreshSlideImage();
  slideImgTimer = setInterval(refreshSlideImage, 2000);
}

function stopSlideImageRefresh() {
  if (slideImgTimer) {
    clearInterval(slideImgTimer);
    slideImgTimer = null;
  }
}

// ── Keyboard shortcuts (no Space — conflicts with PowerPoint) ──
document.addEventListener('keydown', function(e) {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
  switch (e.key) {
    case '+': case '=': changeFontSize(0.3); break;
    case '-': case '_': changeFontSize(-0.3); break;
    case 't': case 'T': toggleTimer(); break;
    case 'm': case 'M': toggleMirror(); break;
    case 'p': case 'P': togglePortrait(); break;
    case 'h': case 'H': toggleHighlight(); break;
    case 'd': case 'D': toggleDemoMode(); break;
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
addTouchButton('wordSpUp', function() { changeWordSpacing(WORDSP_STEP); });
addTouchButton('wordSpDown', function() { changeWordSpacing(-WORDSP_STEP); });
addTouchButton('mirrorBtn', toggleMirror);
addTouchButton('portraitBtn', togglePortrait);
addTouchButton('dispPrev', function() { changeDisplay(-1); });
addTouchButton('dispNext', function() { changeDisplay(1); });
document.getElementById('highlightSlider').addEventListener('input', function() {
  setHighlightLevel(parseInt(this.value));
});
addTouchButton('panelPrevSlide', function() {
  fetch('/api/ppt/prev').catch(function(){});
});
addTouchButton('panelNextSlide', function() {
  fetch('/api/ppt/next').catch(function(){});
});
addTouchButton('demoToggleBtn', toggleDemoMode);

// ── Slide panel drag-to-resize (portrait mode) ──
(function() {
  var drag = document.getElementById('slidePanelDrag');
  var panel = document.getElementById('slidePanel');
  var dragging = false;
  var startY = 0;
  var startH = 0;
  var MIN_H = 80;
  var MAX_RATIO = 0.7;  // max 70% of window height

  function getY(e) {
    if (e.touches && e.touches.length) return e.touches[0].clientY;
    return e.clientY;
  }

  function onStart(e) {
    e.preventDefault();
    dragging = true;
    startY = getY(e);
    startH = panel.offsetHeight;
    document.body.style.cursor = 'ns-resize';
    document.body.style.userSelect = 'none';
    document.body.style.webkitUserSelect = 'none';
  }

  function onMove(e) {
    if (!dragging) return;
    e.preventDefault();
    var delta = startY - getY(e);  // dragging up = positive = bigger panel
    var newH = Math.max(MIN_H, Math.min(window.innerHeight * MAX_RATIO, startH + delta));
    panel.style.height = newH + 'px';
    // Update the image max-height to fill available space
    // Subtract roughly 100px for nav buttons, info text, drag handle
    var imgMaxH = Math.max(40, newH - 100);
    document.documentElement.style.setProperty('--panel-img-max-h', imgMaxH + 'px');
  }

  function onEnd() {
    if (!dragging) return;
    dragging = false;
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
    document.body.style.webkitUserSelect = '';
    // Save the panel height as a percentage of window for persistence
    var pct = Math.round((panel.offsetHeight / window.innerHeight) * 100);
    fetch('/api/settings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ panelHeightPct: pct })
    }).catch(function(){});
  }

  drag.addEventListener('mousedown', onStart);
  drag.addEventListener('touchstart', onStart, { passive: false });
  document.addEventListener('mousemove', onMove);
  document.addEventListener('touchmove', onMove, { passive: false });
  document.addEventListener('mouseup', onEnd);
  document.addEventListener('touchend', onEnd);

  // Restore saved panel height
  fetch('/api/settings', { method: 'GET' }).catch(function(){});
})();

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
    parser.add_argument("--public-url",
                        help="Public URL for phone remote (e.g. ngrok URL). "
                             "Used for QR code and remote link instead of local IP.")

    args = parser.parse_args()

    global script_sections, total_slides, mode, slideshow_active
    global local_ip, server_port, deck_path, public_url
    global demo_script, demo_dir

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
    public_url = args.public_url or cfg.get("public_url")

    server_port = port

    # ── Parse script ──
    script_path = Path(script_file)
    if not script_path.exists():
        print(f"Error: Script file not found: {script_file}")
        sys.exit(1)

    script_sections, slide_titles = parse_script(str(script_path))
    total_slides = len(script_sections)
    print(f"Loaded {total_slides} script sections from {script_path.name}")

    # ── Load persistent settings ──
    global _settings_file_path
    talk_dir = script_path.parent
    _settings_file_path = str(talk_dir / ".teleprompter-settings.json")
    _load_saved_settings()

    # Apply saved screenshot display if present
    saved_disp = display_settings.get("screenshotDisplay")
    if saved_disp:
        _set_screenshot_display(saved_disp if saved_disp != 0 else None)

    # ── Auto-detect demo script in the same directory ──
    demo_files = sorted(talk_dir.glob("demo-*.py"))
    if demo_files:
        demo_script = str(demo_files[0])
        demo_dir = str(talk_dir)
        print(f"Demo script: {demo_files[0].name}")
    else:
        print("No demo script found (looked for demo-*.py)")

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

    # ── Live screenshot capture loop ──
    if platform.system() == "Darwin":
        print("Live slide preview enabled (captures PowerPoint window every 2s)")
        print("  Note: macOS may require Screen Recording permission for Terminal/Python")
        threading.Thread(target=_run_screenshot_loop, daemon=True).start()

    # ── Start HTTP server ──
    server = http.server.HTTPServer(("0.0.0.0", port), TeleprompterHandler)

    local_ip = _detect_local_ip() or "YOUR_IP"
    if local_ip == "YOUR_IP":
        print("  Warning: Could not detect LAN IP. Replace YOUR_IP with your computer's IP.")

    print(f"\nTeleprompter running at http://localhost:{port}")
    if public_url:
        print(f"\n  Phone remote:  {public_url.rstrip('/')}/remote  (public)")
    else:
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

    # ── Background slideshow status checker ──
    # Periodically checks if PowerPoint slideshow is still running
    # so the UI updates even without the VBA callback
    if platform.system() == "Darwin":
        def _check_slideshow_status():
            import time as _time
            while True:
                _time.sleep(3)
                if slideshow_active:
                    ok, result = _ppt_applescript("check")
                    if ok and result == "stopped":
                        _receive_stopped()
                        print("  Slideshow ended (detected by status check)")
        threading.Thread(target=_check_slideshow_status, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        server.shutdown()


if __name__ == "__main__":
    main()
