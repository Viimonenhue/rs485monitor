RS-485 Serial Monitor — build instructions
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

Using the app
--------------
- Left panel: pick COM port, baud rate, data bits, parity, stop bits
  (matches the fields in your current tool), then "Open Port".
- "Live Log" tab: scrolling timestamped hex log, auto-scrolls, shows a
  running received count.
- "Live Graph" tab: tick the checkboxes for whichever field(s) you want
  plotted live. Labels now match your Excel column mapping — Enc M1,
  Enc M2, Effekt kW M1, Effekt kW M2, Strom A M1, Strom A M2, and
  Felkod (fault code) are named; the remaining unspecified fields show
  as "Field 7" .. "Field 11". "Effekt kW M1" is ticked by default. The
  time axis is in real seconds, using your device's fixed 0.02s
  sample interval (not PC arrival time).
- Any row where Felkod (the last field) is non-zero is highlighted red
  in the Live Log so faults jump out immediately.
- "Save as .txt" / "Save as .xlsx": dumps everything received so far,
  with a timestamp column, in your exact original hex format.
- "Auto-save to .txt while running": streams every line to a file as
  it arrives, so nothing is lost if the app crashes or you forget to
  save.

Notes
-----
- The app treats incoming serial data as newline-terminated ASCII text
  (matching your log's format: "00FB;02D0;0019;...;E612\n"). If your
  adapter actually sends raw binary instead of ASCII hex text, tell me
  and I'll switch the parser to read raw bytes and hex-encode them
  instead.
- RTS/DTR checkboxes are wired to the serial port's control lines, same
  as your current tool.
- Once you tell me more about the graph (which field is voltage/current/
  status/etc., and whether you want multiple traces with different
  scales), I can add labels, units, and thresholds/alarm lines.
