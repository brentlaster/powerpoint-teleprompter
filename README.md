# Slide Teleprompter

A Python-based teleprompter that syncs with your PowerPoint slideshow in real time. As you advance slides, the corresponding section of your speaker script automatically appears in a browser window on your second monitor. Includes a phone-based remote control so you can adjust the display and advance slides without losing PowerPoint focus.

## How It Works

```
PowerPoint (VBA macro)
    |
    |  HTTP: /api/slide/{index}/{total}
    v
Python server (0.0.0.0:8765)
    |            |
    |  polling   |  settings + scroll commands
    v            v
Browser         Phone Remote
teleprompter    (http://your-ip:8765/remote)
```

1. A small **VBA macro** inside your PowerPoint detects every slide change and sends an HTTP request to a local Python server.
2. The **Python server** (`teleprompter.py`) tracks the active slide and manages display settings.
3. A **browser-based teleprompter UI** polls the server and displays the matching script section.
4. A **phone remote control** lets you adjust the teleprompter, scroll the text, and advance slides -- all without touching the teleprompter screen or losing PowerPoint focus.

This architecture avoids PowerPoint for Mac's sandbox restrictions -- no file I/O permissions needed, no AppleScript quirks. The VBA macro simply runs `curl` in the background on each slide transition.

## Features

- **Real-time sync** -- VBA macro notifies the teleprompter on every slide change
- **Phone remote control** -- adjust display, scroll text, and advance slides from your phone
- **Slide control buttons** -- Prev/Next Slide and Start Slideshow buttons on both the main page Controls panel and the phone remote, so you can advance slides without touching PowerPoint
- **Portrait mode** -- optimized vertical layout with a live slide preview panel at the bottom, ideal for a portrait-oriented second monitor
- **Live slide preview** (macOS) -- captures a screenshot of the display running PowerPoint every 2 seconds (and immediately on slide change) for the portrait mode preview panel; includes a display picker for multi-monitor setups
- **Highlight bar** -- adjustable-opacity yellow highlight band at the top of the teleprompter to mark your reading position; controlled via slider on the Controls panel or phone remote
- **Scroll progress indicator** -- a percentage and down-arrows on the left edge show how far through the current slide's script you've read, with one arrow per remaining screen of text; only appears when text overflows the screen
- **Scroll past end** -- text can be scrolled completely off the top of the screen, so you're never stuck with text anchored at the bottom
- **Demo Mode** (macOS) -- one-tap switch between your PowerPoint slideshow and a Terminal window for live coding demos, then back again on the same slide
- **Expandable Q&A index** -- add a collapsible Q&A section to your Questions slide using HTML `<details>`/`<summary>` tags in your script
- **Mirror mode** -- horizontal flip for physical teleprompter rigs
- **Config file & launcher** -- point to your deck + script in a JSON file and launch everything with one command

## Requirements

- **Python 3.6+** (no additional packages for VBA mode)
- **PowerPoint for Mac** (or Windows -- the VBA macro works on both)
- **A web browser** for the teleprompter display
- **A phone** on the same Wi-Fi network (for the remote control)
- **Optional:** `pynput` for keyboard fallback mode (`pip3 install pynput`)

## Quick Start

### 1. Start the server

```bash
python3 teleprompter.py your-script.md
```

The terminal will print:
- The VBA macro to paste into PowerPoint
- The teleprompter URL (`http://localhost:8765`)
- The phone remote URL (`http://your-ip:8765/remote`)
- The QR code page URL (`http://localhost:8765/qr`)

### 2. Insert the VBA macro into PowerPoint

1. Open your `.pptx` in PowerPoint.
2. Go to **Tools > Macro > Visual Basic Editor** (or `Opt+F11` / `Alt+F11`).
3. Click **Insert > Module**.
4. Paste the macro code from the terminal output (also saved to `vba_macro.txt`).
5. Close the VBA editor.
6. Optionally save as `.pptm` (macro-enabled) so the macro persists.

### 3. Set up the phone remote

When the teleprompter opens in your browser, expand the **Controls** panel (bottom-right) where you'll find a QR code at the top. Scan it with your phone camera to open the remote control page. Your phone must be on the same Wi-Fi network as your computer. A dedicated QR page is also available at `http://localhost:8765/qr`.

**macOS firewall note:** The first time you run the server, macOS may ask to allow Python to accept incoming connections. Click **Allow**. If the phone can't connect, check **System Settings > Network > Firewall** and ensure Python is permitted.

### 4. Arrange your screens and present

- **Screen 1 (audience-facing):** PowerPoint slideshow
- **Screen 2 (speaker-facing):** Browser at `http://localhost:8765`
- **Your phone:** Remote control

Start the slideshow. The teleprompter automatically displays the matching script section as you advance slides.

## Script File Format

Your speaker script should be a Markdown file with `SLIDE` markers separating sections:

```markdown
## SLIDE 1: TITLE

Good morning, everyone. Welcome to today's talk...

## SLIDE 2: AGENDA

Here's what we'll cover today...

## SLIDE 3: FIRST TOPIC

Let's dive into the first topic...
```

Supported marker formats include `## SLIDE 1`, `## [SLIDE 1 -- Title]`, `**SLIDE 3**`, `SLIDE 11b`, and similar. The key requirement is that the word `SLIDE` followed by a number appears on its own line. The parser maps markers 1:1 in order to slides in the deck.

## Phone Remote Control

The phone remote (`/remote`) is the primary way to control the teleprompter during a presentation. It communicates with the server over HTTP, so it never steals focus from PowerPoint.

### Slide Control

Green **Prev Slide** and **Next Slide** buttons tell PowerPoint to change slides via AppleScript (macOS). Your clicker continues to work normally alongside these.

**macOS permission:** The first time you use these buttons, macOS may ask you to grant Python permission to control PowerPoint under **System Settings > Privacy & Security > Automation**. Allow it.

### Scroll Control

Blue **Scroll Up** and **Scroll Down** buttons scroll the teleprompter text. Each tap scrolls approximately 300 pixels with smooth animation.

### Demo Mode (macOS)

A large toggle button switches between PowerPoint and Terminal for live coding demos. When you tap **Switch to Demo**, the teleprompter opens a Terminal window over the running slideshow, `cd`s into the talk directory, and types the demo command (without executing it) so you can hit Return when ready. Tap **Back to Slides** to stop the demo process, close Terminal, and return to the slideshow on the exact slide where you left off. The keyboard shortcut `D` also toggles demo mode.

Demo mode is available when a `demo-*.py` file exists in the same directory as your speaker script. The teleprompter auto-detects it at startup. If no demo script is found, the button is hidden.

### Display Settings

All display changes made on the phone are applied to the teleprompter in real time:

| Control | Range | Default | Description |
|---------|-------|---------|-------------|
| **Font** | 0.8 -- 6.0 | 2.8 | Font size in rem units |
| **Width** | 600 -- 3000px | 1800px | Max width of the text column |
| **Spacing** | -4 -- 24px | 0px | Extra space between words |
| **Highlight Bar** | 0 -- 100% | Off (0%) | Yellow highlight band at top of screen; slide to adjust opacity |
| **Mirror** | On / Off | Off | Horizontal flip for physical teleprompter rigs |

## Teleprompter Window Controls

The teleprompter browser window has a collapsible floating control panel (bottom-right) with all the same controls as the phone remote, plus a scannable QR code for quick phone setup and Prev/Next Slide buttons. The Controls panel is scrollable if your window is too short to show everything. Changes made here sync back to the phone. All buttons are touch-friendly (60px minimum) in case you use a touchscreen monitor.

**Note:** Touching the teleprompter screen may steal focus from PowerPoint and cause your clicker to stop working. Use the phone remote instead to avoid this issue. See the Touchscreen Focus section below for details.

### Portrait Mode

Toggle Portrait mode from the Controls panel or press `P`. This switches the teleprompter to a vertical layout optimized for portrait-oriented monitors. A slide info panel appears at the bottom showing the current slide number/title, Prev/Next buttons, and a live screenshot preview of the PowerPoint slideshow (macOS only).

The live preview captures the display running PowerPoint every 2 seconds and immediately on every slide change. It requires macOS **Screen Recording** permission for Terminal or Python (System Settings > Privacy & Security > Screen Recording).

On multi-monitor setups, the **Capture** control in the Controls panel lets you select which display to capture. Use the left/right arrows to cycle through Auto (tries display 2 first), Display 1, Display 2, Display 3, etc. The display number is cached for the duration of the slideshow and resets when the slideshow ends.

### Scroll Progress Indicator

When a slide has more script text than fits on screen, an indicator appears on the left edge showing your scroll progress (e.g., "0%" at the top, "100%" at the bottom). Down arrows appear below the percentage — one arrow per remaining screen of text. For example, if there are 3 screens of text, you'll see 3 arrows initially; scrolling past the first screen reduces it to 2 arrows, and so on until you reach the bottom. The indicator is hidden entirely when all text fits on screen without scrolling.

### Highlight Bar

The highlight bar places a yellow semi-transparent band over the first few lines of text in the teleprompter area. It tracks your scroll position so the highlighted region always covers the topmost visible text. Use the slider in the Controls panel (or phone remote) to adjust the intensity from 0% (off) to 100% (full opacity). Press `H` to toggle it on/off via keyboard (toggles between 0% and 50%).

### Keyboard Shortcuts

These work when the teleprompter browser window has focus:

| Key | Action |
|-----|--------|
| `+` / `=` | Increase font size |
| `-` / `_` | Decrease font size |
| `D` | Toggle demo mode (switch between slides and Terminal) |
| `P` | Toggle portrait mode |
| `H` | Toggle highlight bar (0% / 50%) |
| `T` | Start / pause timer |
| `M` | Toggle mirror mode |

## Expandable Q&A Index

You can add an expandable Q&A section to your Questions slide by including HTML `<details>` and `<summary>` tags directly in your script file. The teleprompter renders these as collapsible entries -- tap a question to reveal the answer.

```markdown
## SLIDE 47: QUESTIONS

Thank you! I'd love to take your questions.

<div class="qa-index">

<details>
<summary>What's the difference between X and Y?</summary>
<div class="qa-answer">
X handles ... while Y focuses on ...
</div>
</details>

<details>
<summary>How does this scale in production?</summary>
<div class="qa-answer">
In production you would ...
</div>
</details>

</div>
```

The Q&A index is styled with indented borders, a triangle expand indicator, and a larger answer font for readability. Questions stay collapsed by default so the slide isn't cluttered, and you can quickly open any one during the Q&A session.

## Command-Line Options

```bash
# Basic usage (VBA mode)
python3 teleprompter.py path/to/script.md

# Custom port (macro must be re-pasted if port changes)
python3 teleprompter.py path/to/script.md --port 9000

# Don't auto-open browser
python3 teleprompter.py path/to/script.md --no-browser

# Keyboard fallback mode (no VBA needed)
pip3 install pynput  # one-time
python3 teleprompter.py path/to/script.md --keyboard
```

### Keyboard Fallback Mode

If you can't use VBA (e.g., presenting from someone else's machine), keyboard mode lets you advance the teleprompter script manually:

| Key | Action |
|-----|--------|
| `.` (period) | Next script section |
| `,` (comma) | Previous script section |
| Home | Jump to first section |
| End | Jump to last section |

PowerPoint's arrow keys and clicker continue to control slides independently.

## Config File and Launcher

For a streamlined workflow, you can create a JSON config file that points to your deck and script and then launch everything with a single command.

### Config File Format

Create a JSON file (e.g., `talk.json`) in your talk directory:

```json
{
    "deck": "my-talk.pptm",
    "script": "my-talk_script.md",
    "port": 8765,
    "auto_start": false
}
```

| Field | Required | Default | Description |
|-------|----------|---------|-------------|
| `script` | **Yes** | -- | Path to the speaker script `.md` file |
| `deck` | No | -- | Path to the `.pptm` (macro-enabled) PowerPoint file |
| `port` | No | 8765 | Server port |
| `auto_start` | No | false | Automatically start the slideshow after opening the deck |

Paths can be absolute or relative to the config file's directory. For example, if `talk.json` is in `~/talks/my-talk/`, then `"script": "my-talk_script.md"` resolves to `~/talks/my-talk/my-talk_script.md`.

### Using the Config File

```bash
python3 teleprompter.py --config talk.json
```

This single command will:

1. Open the PowerPoint deck (if `deck` is specified)
2. Start the teleprompter server
3. Open the teleprompter in your browser (QR code is in the Controls panel)
4. Optionally start the slideshow automatically (if `auto_start` is true)

You can also override the deck path on the command line:

```bash
python3 teleprompter.py --config talk.json --deck other-version.pptm
```

### Launcher Script

A convenience shell script (`launch.sh`) is included:

```bash
# Run from your talk directory (picks up ./talk.json automatically)
cd ~/talks/keynote
/path/to/launch.sh

# Or specify a config file explicitly
/path/to/launch.sh ~/talks/keynote/talk.json
```

### Typical Workflow with Config

1. Create your deck (`my-talk.pptm`) and script (`my-talk_script.md`) in a folder.
2. Create a `talk.json` config file in the same folder.
3. `cd` into the folder and run `launch.sh`. It finds `talk.json` in the current directory, starts the server, and opens the teleprompter in your browser.
4. Expand the Controls panel and scan the QR code with your phone.
5. Tap **Start Slideshow** on the phone remote when ready.
6. Present using your clicker for slides and the phone remote for scrolling.

## Server Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/` | GET | Teleprompter UI |
| `/remote` | GET | Phone remote control UI |
| `/qr` | GET | QR code page for phone remote URL |
| `/api/state` | GET | Current slide, script HTML, settings, scroll commands |
| `/api/settings` | GET/POST | Read or update display settings |
| `/api/slide/{idx}/{total}` | GET | Called by VBA macro on slide change |
| `/api/stopped` | GET | Called by VBA macro when slideshow ends |
| `/api/goto/{n}` | GET | Jump to slide number |
| `/api/ppt/next` | GET | Advance PowerPoint slide (macOS, via AppleScript) |
| `/api/ppt/prev` | GET | Go back one PowerPoint slide (macOS, via AppleScript) |
| `/api/scroll/up` | GET | Scroll teleprompter text up |
| `/api/scroll/down` | GET | Scroll teleprompter text down |
| `/api/scroll/top` | GET | Scroll teleprompter text to top |
| `/api/ppt/start` | GET | Start PowerPoint slideshow (macOS, via AppleScript) |
| `/api/ppt/stop` | GET | Stop PowerPoint slideshow (macOS, via AppleScript) |
| `/api/slide-image` | GET | Live screenshot of the PowerPoint slideshow (JPEG) |
| `/api/slide-image-debug` | GET | Debug info for the screenshot pipeline |
| `/api/screenshot-display` | POST | Set which display to capture (JSON: `{"display": 3}`, 0 = auto) |
| `/api/demo/toggle` | GET | Toggle demo mode (switch between slideshow and Terminal) |
| `/api/demo/state` | GET | Current demo mode state and available demo script |
| `/api/remote-url` | GET | Get the phone remote URL with LAN IP |

## File Structure

```
my-talk/
  my-talk.pptx              # Original slide deck
  my-talk.pptm              # Macro-enabled version (after setup)
  my-talk_script.md          # Speaker script with SLIDE markers
  demo-example.py            # Live demo script (auto-detected by teleprompter)
  talk.json                  # Config file (deck + script paths, settings)
  vba_macro.txt              # Auto-generated macro (created by teleprompter.py)
  teleprompter.py            # The teleprompter server
  launch.sh                  # Convenience launcher script
```

## Troubleshooting

**"Waiting for slideshow..." never changes**
The VBA macro isn't firing. Verify you pasted it into a Module (not ThisPresentation or a Class), and that macros are enabled in PowerPoint's security settings (Preferences > Security & Privacy > Enable all macros).

**Script shows but doesn't change per slide**
Restart the Python server. You should see `Loaded N script sections` on startup where N matches your slide count.

**Wrong section for a slide**
The parser maps SLIDE markers 1:1 in order. Ensure your script and deck have the same number of slides.

**Phone can't connect to remote**
Check that your phone and computer are on the same Wi-Fi network. On macOS, ensure the firewall allows Python to accept incoming connections (System Settings > Network > Firewall). Verify the IP address printed in the terminal matches your computer's actual LAN IP (`ifconfig | grep "inet "` on the en0 interface).

**Phone can't connect on hotel/conference WiFi**
Many hotel and conference networks use client isolation, which prevents devices on the same WiFi from communicating with each other. The easiest workaround is to use your **phone as a mobile hotspot**: enable the hotspot on your phone, connect your Mac to it, then open the remote URL on the phone's browser. Both devices share the phone's network, so the connection works. Alternatively, use **Mac Internet Sharing** (System Settings > General > Sharing > Internet Sharing) to create a local WiFi hotspot from your Mac that the phone can join.

**Slide control buttons on phone don't work**
macOS needs permission for Python to control PowerPoint. Go to System Settings > Privacy & Security > Automation and allow Python (or Terminal) to control Microsoft PowerPoint.

**PowerPoint prompts about macros every time**
Go to PowerPoint > Preferences > Security & Privacy and set "Enable all macros" to suppress the prompt. This applies to all presentations, so only enable it if you trust the files you open.

**Demo mode button doesn't appear**
The teleprompter looks for a file matching `demo-*.py` in the same directory as your speaker script. If none is found, the demo toggle is hidden. Check that your demo script follows the naming convention (e.g., `demo-example.py`).

**Demo mode: Terminal doesn't appear over slideshow**
macOS needs permission for Python to control Terminal under System Settings > Privacy & Security > Automation. Allow Python (or Terminal) to send events to other applications.

**Portrait mode: no slide preview image**
The live screenshot feature requires macOS Screen Recording permission. Go to System Settings > Privacy & Security > Screen Recording and ensure Terminal (or Python) is listed and enabled. You may need to restart Terminal after granting the permission. Visit `http://localhost:8765/api/slide-image-debug` while the slideshow is running to see diagnostic info about the screenshot pipeline.

**Portrait mode: screenshot shows wrong display**
By default, the teleprompter auto-detects and captures display 2 (assuming the teleprompter is on display 1). On setups with 3+ monitors, it may pick the wrong one. Use the **Capture** control in the Controls panel to cycle through displays until you see your PowerPoint slideshow. The selected display is cached for the rest of the slideshow session.

**Start Slideshow button doesn't reappear when slideshow ends**
A background thread checks PowerPoint's slideshow status every 3 seconds. If it detects the slideshow has ended, it resets the UI. Make sure PowerPoint is still running; the detection relies on AppleScript checking the slide show window count.

**Touchscreen focus issues**
Touching the teleprompter browser on a secondary touchscreen may not move macOS keyboard focus away from the primary screen. This is a macOS limitation -- the OS does not automatically switch focus between displays on touch input alone. Use the phone remote for all controls during a presentation to avoid focus conflicts with PowerPoint.

**Port already in use**
Another instance may be running. Kill it with `lsof -ti:8765 | xargs kill` or use `--port` to pick a different port.

**Keyboard mode: keys not detected**
macOS requires accessibility permissions for `pynput`. Go to System Settings > Privacy & Security > Accessibility and grant permission to your terminal app.
