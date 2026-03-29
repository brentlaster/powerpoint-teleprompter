# Slide Teleprompter

A Python-based teleprompter that syncs with your PowerPoint slideshow in real time. As you advance slides (including through animations), the corresponding section of your speaker script automatically appears in a browser window on your second monitor.

## How It Works

The system has three parts working together:

1. **A small VBA macro** inside your PowerPoint presentation detects every slide change and sends an HTTP request to a local server.
2. **A Python HTTP server** (`teleprompter.py`) receives those requests and tracks which slide is active.
3. **A browser-based teleprompter UI** polls the server every 200ms and displays the matching section of your speaker script.

This architecture avoids all of PowerPoint for Mac's sandbox restrictions — no file I/O permissions needed, no AppleScript quirks. The VBA macro simply runs `curl` in the background on each slide transition.

```
PowerPoint (VBA macro)
    │
    │  HTTP: /api/slide/{index}/{total}
    ▼
Python server (localhost:8765)
    │
    │  JSON polling: /api/state
    ▼
Browser teleprompter UI
```

## Requirements

- **Python 3.6+** (no additional packages needed for VBA mode)
- **PowerPoint for Mac** (or Windows — the VBA macro works on both)
- **A web browser** for the teleprompter display
- **Optional:** `pynput` package for keyboard fallback mode (`pip3 install pynput`)

## Script File Format

Your speaker script should be a Markdown file with `SLIDE` markers separating sections. Each marker corresponds to one slide in the deck. The parser handles several common formats:

```markdown
## SLIDE 1: TITLE

Good morning, everyone. Welcome to today's talk...

[TIMING: 2:00]

---

## SLIDE 2: AGENDA

Here's what we'll cover today...

[TIMING: 3:30]

---

## SLIDE 3: FIRST TOPIC

Let's dive into the first topic...
```

Supported marker variations include `## SLIDE 1`, `## [SLIDE 1 — Title]`, `**SLIDE 3**`, `SLIDE 11b`, and similar patterns. The key requirement is that the word `SLIDE` followed by a number appears on its own line.

## Setup (One-Time Per Presentation)

### 1. Start the teleprompter server

```bash
cd /path/to/your/talks
python3 teleprompter.py your-script_with_cues.md
```

The server prints the VBA macro code to your terminal and also saves it to `vba_macro.txt` in the same directory as your script.

### 2. Insert the VBA macro into PowerPoint

1. Open your `.pptx` in PowerPoint.
2. Go to **Tools → Macro → Visual Basic Editor** (or press `Opt+F11`).
3. In the VBA editor, click **Insert → Module**.
4. Paste the macro code (from the terminal output or `vba_macro.txt`).
5. Close the VBA editor.

### 3. Save as macro-enabled (recommended)

To avoid repeating step 2 every time, save the presentation as `.pptm` (macro-enabled format) via **File → Save As**. The macro will persist across sessions.

### 4. Arrange your screens

- **Screen 1 (audience-facing):** PowerPoint slideshow
- **Screen 2 (speaker-facing):** Browser window at `http://localhost:8765`

### 5. Start the slideshow

The teleprompter will automatically begin displaying script sections as you advance through slides. Animations work normally — the script only changes when you move to a new slide, not on animation clicks.

## Usage

### Basic (VBA mode — recommended)

```bash
python3 teleprompter.py path/to/script.md
```

### Custom port

```bash
python3 teleprompter.py path/to/script.md --port 9000
```

If you change the port, you'll need to re-paste the macro (it embeds the port number).

### Without auto-opening browser

```bash
python3 teleprompter.py path/to/script.md --no-browser
```

### Keyboard fallback mode

If you can't use VBA (e.g., presenting from someone else's machine), use keyboard mode instead:

```bash
pip3 install pynput  # one-time install
python3 teleprompter.py path/to/script.md --keyboard
```

In this mode, you advance the teleprompter manually:

- **Period (`.`)** — next script section
- **Comma (`,`)** — previous script section
- **Home** — jump to first section
- **End** — jump to last section

Use period/comma to advance the script independently of PowerPoint's arrow keys / clicker, which continue to control slides and animations as normal.

## Teleprompter UI Controls

The browser UI includes several controls:

| Control | Action |
|---|---|
| `+` / `=` key | Increase font size |
| `-` / `_` key | Decrease font size |
| `T` key | Start/pause timer |
| `M` key | Toggle mirror mode (for physical teleprompter rigs) |
| **A+** / **A−** buttons | Font size (top bar) |
| **Mirror** checkbox | Mirror mode (top bar) |

The top bar also shows the current slide number, total slides, a progress bar, and an elapsed timer that starts automatically when the slideshow begins.

## File Structure

For each talk, you'll typically have:

```
my-talk/
├── my-talk.pptx              # Original slide deck
├── my-talk.pptm              # Macro-enabled version (after setup)
├── my-talk_with_cues.md       # Speaker script with SLIDE markers
└── vba_macro.txt              # Auto-generated macro (created by teleprompter.py)
```

## Troubleshooting

**"Waiting for slideshow..." never changes**
The VBA macro isn't firing. Verify you pasted it into a Module (not ThisPresentation or a Class), and that macros are enabled in PowerPoint's security settings (Preferences → Security & Privacy → Enable all macros).

**Script shows but doesn't change per slide**
Restart the Python server — it may be running an older version of the code. You should see `Loaded N script sections` on startup where N matches your slide count.

**Wrong section for a slide**
The parser maps SLIDE markers 1:1 in order. If your script has 39 SLIDE markers, slide 1 maps to the first section, slide 2 to the second, etc. Ensure your script and deck have the same number of slides.

**Port already in use**
Another instance may be running. Kill it with `lsof -ti:8765 | xargs kill` or use `--port` to pick a different port.

**Keyboard mode: keys not detected**
macOS requires accessibility permissions for `pynput`. Go to System Settings → Privacy & Security → Accessibility and grant permission to your terminal app.
