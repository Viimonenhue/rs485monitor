RS-485 Serial Monitor
============================================

What this is
------------
A Windows desktop tool for connecting to an RS-485-to-USB adapter and
monitoring hex-formatted telemetry from a gate/motor automation
controller in real time — live table, live multi-parameter graph,
threshold overlays, and export to file, all in one app. No data sent
back to the device; this is a passive listener.

The device sends one semicolon-separated hex line per sample, e.g.:

    0058;02D0;0068;0000;12;00;01;00;01;00;02;E932

Each field maps to a named, unit-aware value:

    Field        Code    Meaning              Conversion
    -----------  ------  -------------------  --------------------------
    Encoder 1    L111    Position (deg)       raw = degrees, wraps 0-359
    Encoder 2    L121    Position (deg)       raw = degrees, wraps 0-359
    Effekt kW M1 C231    Motor 1 power (kW)    raw / 100, capped 1.99 kW
    Effekt kW M2 C241    Motor 2 power (kW)    raw / 100, capped 1.99 kW
    Strom A M1   C251    Motor 1 current (A)   raw / 10, capped 20 A
    Strom A M2   C261    Motor 2 current (A)   raw / 10, capped 20 A
    C211 M1/M2, C212 M1/M2, Port B - unnamed status/reserved fields
    Felkod               Fault code            non-zero = row highlighted

Main features
-------------
- Live Log tab: table view (named columns, tightened widths, hex value
  + real-world value shown together) or a raw Hex view toggle that
  matches the original Serial Debug Assistant output exactly. Screen
  repaints in configurable batches (Settings menu) so the ~50 lines/sec
  data rate never overwhelms the UI - nothing is ever dropped, this
  only controls how often the table redraws.
- Live Graph tab: tick any combination of fields to plot. Each gets its
  own color-coded Y-axis, all sharing one aligned grid (same tick count,
  evenly spaced per axis, so gridlines actually line up across very
  different ranges like degrees vs kW). Y-axis floors at 0 for physical
  units. Pan/zoom toolbar with a Reset Zoom button; your zoomed view
  survives live updates until you reset it.
- Threshold lines (all optional, all just visual reference lines - not
  computed from live logic): Belastningsvakt (load guard) and
  Motorskydd (motor protection) are per-motor and per-direction
  (Öppna/dashed, Stäng/dotted). Personskydd is a single global setting,
  shown as dotted lines at +0.25/-0.10 kW. Each control only appears
  when its matching field is ticked.
- Saved Tests: save the current run's raw hex under a name (persists to
  a "saved_tests" folder next to the .exe), then overlay it on the live
  graph later for comparison. Overlay traces are lighter/dashed so
  they're clearly reference data, not live.
- Export: .txt, .xlsx (raw hex), .xlsx (hex + computed values), graph
  image (PNG/PDF/SVG), or "Export All" to write all four at once from a
  single filename prompt. Auto-save streams every line to a .txt file
  as it arrives, so a crash or forgotten save never loses data.

Notes
-----
- Incoming serial data is treated as newline-terminated ASCII text.
  Stray NUL bytes (a known RS-485 half-duplex artifact) are stripped on
  read, since Tk silently truncates any text containing one.
- If a field comes through blank on a given line, the table/graph carry
  forward the last known value for it (the saved raw file always keeps
  the exact original data, blanks included).
- RTS/DTR checkboxes are wired to the serial port's control lines.


Build instructions
============================================

Two ways to get the Windows .exe. Pick whichever's easier for you.

OPTION A — Build automatically with GitHub Actions (no Python needed)
-----------------------------------------------------------------------
1) Create a new repo on GitHub, e.g. "rs485-monitor".
2) Push everything in this folder to it (including the hidden
   .github/workflows/build.yml file — that's the important one):

     git init
     git add .
     git commit -m "Initial commit"
     git branch -M main
     git remote add origin https://github.com/YOUR-USERNAME/rs485-monitor.git
     git push -u origin main

3) On GitHub, open the "Actions" tab of your repo. A workflow called
   "Build Windows EXE" will already be running (it triggers on every
   push to main). Wait for the green checkmark.
4) Click into that run, scroll to "Artifacts", and download
   "RS485Monitor-windows" — it's a zip containing RS485Monitor.exe.

Every time you push a change, a fresh .exe gets built automatically —
no local Python install, no PyInstaller, nothing to set up on your PC.

OPTION B — Build locally on your own PC
-----------------------------------------
1) Install Python 3.10+ from python.org if you don't have it
   (tick "Add python.exe to PATH" during install).

2) Open a Command Prompt in this folder and run:

     pip install -r requirements.txt

3) Build the .exe:

     pyinstaller --onefile --windowed --name RS485Monitor rs485_monitor.py

4) Your standalone program will be at:

     dist\RS485Monitor.exe

   Copy that file wherever you like — it needs no install, no Python,
   nothing else on the target PC.
