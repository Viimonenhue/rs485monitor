name: Build Windows EXE

# Runs on every push to main, and can also be triggered manually from the
# Actions tab ("Run workflow" button).
on:
  push:
    branches: [ "main" ]
  workflow_dispatch: {}

jobs:
  build:
    runs-on: windows-latest

    steps:
      - name: Check out code
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt

      - name: Build EXE with PyInstaller
        run: |
          pyinstaller --onefile --windowed --name RS485Monitor rs485_monitor.py

      - name: Upload EXE as build artifact
        uses: actions/upload-artifact@v4
        with:
          name: RS485Monitor-windows
          path: dist/RS485Monitor.exe
