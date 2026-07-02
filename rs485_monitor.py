"""
RS-485 / USB Serial Monitor
---------------------------
A lightweight Windows desktop tool to connect to an RS-485-to-USB adapter,
show incoming hex-formatted lines live, plot selected fields on a live
graph, and save the raw log to .txt or .xlsx.

Build into a standalone .exe with PyInstaller (see README.txt).
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import time
from datetime import datetime
from collections import deque

import serial
import serial.tools.list_ports

import openpyxl
from openpyxl.utils import get_column_letter

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

APP_TITLE = "RS-485 Serial Monitor"
MAX_LOG_LINES = 5000            # cap the on-screen log so the UI stays fast
GRAPH_REFRESH_MS = 300          # how often the graph redraws
QUEUE_POLL_MS = 50               # how often we drain incoming serial data
SAMPLE_INTERVAL_S = 0.02         # one line = 0.02s (per your device's cycle)

# Named meaning of each semicolon-separated field, in order (A..L in your
# Excel breakdown).
FIELD_NAMES = [
    "Encoder 1",        # A
    "Encoder 2",        # B
    "Effekt kW M1",     # C
    "Effekt kW M2",     # D
    "Strom A M1",       # E
    "Strom A M2",       # F
    "C211 M1",          # G
    "C211 M2",          # H
    "C212 M1",          # I
    "C212 M2",          # J
    "Port B",           # K
    "Felkod",           # L
]


def field_label(index):
    """Human-readable label for field position `index` (0-based)."""
    if index < len(FIELD_NAMES):
        return FIELD_NAMES[index]
    return f"Field {index + 1}"


class SerialMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1150x700")

        self.ser = None
        self.read_thread = None
        self.running = False
        self.data_queue = queue.Queue()

        # Full history of received lines: (timestamp, raw_line, [fields])
        self.raw_rows = []

        # Graph state — unbounded history so the time axis SCALES (zooms
        # out) as more data comes in, rather than scrolling a fixed window.
        self.field_count = 12          # default guess, adjusts to real data
        self.field_vars = []           # BooleanVars, one per field checkbox
        self.field_buffers = []        # list per field, holds ALL values
        self.time_buffer = []          # list of sample times (seconds)
        self.sample_index = 0
        self.graph_paused = False

        # Auto-save state
        self.autosave_path = None
        self.autosave_file = None

        self._build_ui()
        self._refresh_ports()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(QUEUE_POLL_MS, self._poll_queue)
        self.root.after(GRAPH_REFRESH_MS, self._update_graph)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # ---- Left settings panel (always visible) ----
        settings = ttk.Frame(self.root, padding=10)
        settings.pack(side="left", fill="y")

        ttk.Label(settings, text="Connection Settings", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        r = 1
        ttk.Label(settings, text="Port:").grid(row=r, column=0, sticky="w")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(settings, textvariable=self.port_var, width=14, state="readonly")
        self.port_combo.grid(row=r, column=1, sticky="w")
        r += 1
        ttk.Button(settings, text="Refresh Ports", command=self._refresh_ports).grid(
            row=r, column=0, columnspan=2, sticky="we", pady=(2, 8))
        r += 1

        ttk.Label(settings, text="Baud Rate:").grid(row=r, column=0, sticky="w")
        self.baud_var = tk.StringVar(value="125000")
        ttk.Combobox(settings, textvariable=self.baud_var, width=14,
                     values=["9600", "19200", "38400", "57600", "115200", "125000", "230400", "460800"]).grid(
            row=r, column=1, sticky="w")
        r += 1

        ttk.Label(settings, text="Data Bits:").grid(row=r, column=0, sticky="w")
        self.databits_var = tk.StringVar(value="8")
        ttk.Combobox(settings, textvariable=self.databits_var, width=14, state="readonly",
                     values=["5", "6", "7", "8"]).grid(row=r, column=1, sticky="w")
        r += 1

        ttk.Label(settings, text="Parity:").grid(row=r, column=0, sticky="w")
        self.parity_var = tk.StringVar(value="None")
        ttk.Combobox(settings, textvariable=self.parity_var, width=14, state="readonly",
                     values=["None", "Even", "Odd", "Mark", "Space"]).grid(row=r, column=1, sticky="w")
        r += 1

        ttk.Label(settings, text="Stop Bits:").grid(row=r, column=0, sticky="w")
        self.stopbits_var = tk.StringVar(value="1")
        ttk.Combobox(settings, textvariable=self.stopbits_var, width=14, state="readonly",
                     values=["1", "1.5", "2"]).grid(row=r, column=1, sticky="w")
        r += 1

        self.rts_var = tk.BooleanVar(value=True)
        self.dtr_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings, text="RTS", variable=self.rts_var, command=self._apply_line_state).grid(
            row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(settings, text="DTR", variable=self.dtr_var, command=self._apply_line_state).grid(
            row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        self.open_btn = ttk.Button(settings, text="Open Port", command=self._toggle_port)
        self.open_btn.grid(row=r, column=0, columnspan=2, sticky="we", pady=(10, 4))
        r += 1

        self.status_var = tk.StringVar(value="Closed")
        ttk.Label(settings, textvariable=self.status_var, foreground="red").grid(
            row=r, column=0, columnspan=2, sticky="w")
        r += 1

        ttk.Separator(settings, orient="horizontal").grid(row=r, column=0, columnspan=2, sticky="we", pady=10)
        r += 1

        ttk.Label(settings, text="Log / Save", font=("Segoe UI", 11, "bold")).grid(
            row=r, column=0, columnspan=2, sticky="w")
        r += 1

        ttk.Button(settings, text="Save as .txt", command=self._save_txt).grid(
            row=r, column=0, columnspan=2, sticky="we", pady=2)
        r += 1
        ttk.Button(settings, text="Save as .xlsx", command=self._save_xlsx).grid(
            row=r, column=0, columnspan=2, sticky="we", pady=2)
        r += 1
        ttk.Button(settings, text="Clear Log", command=self._clear_log).grid(
            row=r, column=0, columnspan=2, sticky="we", pady=2)
        r += 1

        self.autosave_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings, text="Auto-save to .txt while running",
                         variable=self.autosave_var, command=self._toggle_autosave).grid(
            row=r, column=0, columnspan=2, sticky="w", pady=(8, 0))
        r += 1

        self.count_var = tk.StringVar(value="Received: 0")
        ttk.Label(settings, textvariable=self.count_var).grid(row=r, column=0, columnspan=2, sticky="w", pady=(10, 0))

        # ---- Right side: tabs ----
        notebook = ttk.Notebook(self.root)
        notebook.pack(side="right", fill="both", expand=True, padx=(0, 10), pady=10)

        # --- Tab 1: Live Log ---
        log_tab = ttk.Frame(notebook)
        notebook.add(log_tab, text="Live Log")

        top_row = ttk.Frame(log_tab)
        top_row.pack(anchor="w", fill="x", padx=5, pady=(5, 0))
        self.autoscroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top_row, text="Auto-scroll", variable=self.autoscroll_var).pack(side="left")

        # Header row: names of each field, in order, so the columns of the
        # hex log below are labeled.
        self.log_header_var = tk.StringVar()
        self._refresh_log_header()
        header_label = ttk.Label(log_tab, textvariable=self.log_header_var,
                                  font=("Consolas", 9, "bold"), background="#2b2b2b",
                                  foreground="#ffffff", anchor="w", padding=(6, 4))
        header_label.pack(fill="x", padx=5, pady=(4, 0))

        log_frame = ttk.Frame(log_tab)
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        # wrap="word" so long lines wrap downward and scrolling stays
        # vertical, instead of requiring a horizontal scrollbar.
        self.log_text = tk.Text(log_frame, wrap="word", font=("Consolas", 10))
        yscroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=yscroll.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        # --- Tab 2: Live Graph ---
        graph_tab = ttk.Frame(notebook)
        notebook.add(graph_tab, text="Live Graph")

        controls = ttk.Frame(graph_tab)
        controls.pack(side="top", fill="x", padx=5, pady=5)
        ttk.Label(controls, text="Plot field(s):").pack(side="left")

        self.field_checks_frame = ttk.Frame(controls)
        self.field_checks_frame.pack(side="left", padx=10)
        self._build_field_checkboxes()

        self.pause_btn = ttk.Button(controls, text="Pause Graph", command=self._toggle_pause)
        self.pause_btn.pack(side="right", padx=5)
        ttk.Button(controls, text="Clear Graph", command=self._clear_graph).pack(side="right", padx=5)

        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlabel("Sample #")
        self.ax.set_ylabel("Value (decimal)")
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_tab)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)

    def _build_field_checkboxes(self):
        for w in self.field_checks_frame.winfo_children():
            w.destroy()
        self.field_vars = []
        self.field_buffers = []
        # Default: plot field index 3 (3rd column) checked, rest unchecked -
        # you can tick any you like once you see live data.
        for i in range(self.field_count):
            var = tk.BooleanVar(value=(i == 2))
            cb = ttk.Checkbutton(self.field_checks_frame, text=field_label(i), variable=var)
            cb.pack(side="left", padx=(0, 4))
            self.field_vars.append(var)
            self.field_buffers.append([])
        self._refresh_log_header()

    def _refresh_log_header(self):
        names = [field_label(i) for i in range(getattr(self, "field_count", 12))]
        self.log_header_var.set("   ".join(names) if hasattr(self, "log_header_var") else "")

    # ------------------------------------------------------------- helpers
    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_combo["values"] = ports
        if ports and not self.port_var.get():
            self.port_var.set(ports[0])

    def _parity_code(self):
        return {"None": serial.PARITY_NONE, "Even": serial.PARITY_EVEN,
                "Odd": serial.PARITY_ODD, "Mark": serial.PARITY_MARK,
                "Space": serial.PARITY_SPACE}[self.parity_var.get()]

    def _stopbits_code(self):
        return {"1": serial.STOPBITS_ONE, "1.5": serial.STOPBITS_ONE_POINT_FIVE,
                "2": serial.STOPBITS_TWO}[self.stopbits_var.get()]

    def _apply_line_state(self):
        if self.ser and self.ser.is_open:
            try:
                self.ser.rts = self.rts_var.get()
                self.ser.dtr = self.dtr_var.get()
            except Exception:
                pass

    # ------------------------------------------------------------- connect
    def _toggle_port(self):
        if self.ser and self.ser.is_open:
            self._close_port()
        else:
            self._open_port()

    def _open_port(self):
        port = self.port_var.get()
        if not port:
            messagebox.showwarning(APP_TITLE, "Select a COM port first.")
            return
        try:
            self.ser = serial.Serial(
                port=port,
                baudrate=int(self.baud_var.get()),
                bytesize=int(self.databits_var.get()),
                parity=self._parity_code(),
                stopbits=self._stopbits_code(),
                timeout=0.2,
            )
            self.ser.rts = self.rts_var.get()
            self.ser.dtr = self.dtr_var.get()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not open {port}:\n{e}")
            self.ser = None
            return

        self.running = True
        self.read_thread = threading.Thread(target=self._read_loop, daemon=True)
        self.read_thread.start()
        self.status_var.set(f"Open ({port} @ {self.baud_var.get()})")
        self.open_btn.configure(text="Close Port")

    def _close_port(self):
        self.running = False
        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
        self.ser = None
        self.status_var.set("Closed")
        self.open_btn.configure(text="Open Port")
        self._toggle_autosave_off()

    def _read_loop(self):
        buf = b""
        while self.running and self.ser:
            try:
                chunk = self.ser.read(256)
                if chunk:
                    buf += chunk
                    while b"\n" in buf:
                        line, buf = buf.split(b"\n", 1)
                        text = line.decode(errors="replace").strip("\r").strip()
                        if text:
                            self.data_queue.put(text)
            except Exception:
                break

    # --------------------------------------------------------------- queue
    def _poll_queue(self):
        drained = 0
        while not self.data_queue.empty() and drained < 500:
            line = self.data_queue.get_nowait()
            self._handle_line(line)
            drained += 1
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def _handle_line(self, line):
        ts = datetime.now()
        fields = [f.strip() for f in line.split(";")]

        # Adjust field checkbox count if the device sends a different width
        if len(fields) != self.field_count and len(fields) > 0:
            self.field_count = len(fields)
            self._build_field_checkboxes()

        self.raw_rows.append((ts, line, fields))

        # Fault code is the last field (Felkod). Flag it if non-zero.
        fault = False
        if fields:
            try:
                fault = int(fields[-1], 16) != 0
            except ValueError:
                fault = False

        # ---- Live Log tab ----
        line_start = self.log_text.index("end-1c")
        self.log_text.insert("end", f"{ts.strftime('%H:%M:%S.%f')[:-3]}  {line}\n")
        if fault:
            line_end = self.log_text.index("end-1c")
            self.log_text.tag_add("fault", line_start, line_end)
            self.log_text.tag_configure("fault", foreground="red", font=("Consolas", 10, "bold"))
        # Trim if too long
        num_lines = int(self.log_text.index("end-1c").split(".")[0])
        if num_lines > MAX_LOG_LINES:
            self.log_text.delete("1.0", f"{num_lines - MAX_LOG_LINES}.0")
        if self.autoscroll_var.get():
            self.log_text.see("end")

        self.count_var.set(f"Received: {len(self.raw_rows)}")

        # ---- Graph buffers ----
        # Time axis reflects the device's real 0.02s sample interval, not
        # PC arrival time.
        self.time_buffer.append(self.sample_index * SAMPLE_INTERVAL_S)
        self.sample_index += 1
        for i, f in enumerate(fields):
            if i >= len(self.field_buffers):
                break
            try:
                val = int(f, 16)
            except ValueError:
                val = None
            self.field_buffers[i].append(val)

        # ---- Auto-save ----
        if self.autosave_file:
            self.autosave_file.write(f"{ts.isoformat()}\t{line}\n")
            self.autosave_file.flush()

    # --------------------------------------------------------------- graph
    def _toggle_pause(self):
        self.graph_paused = not self.graph_paused
        self.pause_btn.configure(text="Resume Graph" if self.graph_paused else "Pause Graph")

    def _clear_graph(self):
        self.time_buffer.clear()
        for buf in self.field_buffers:
            buf.clear()
        self.sample_index = 0
        self.ax.clear()
        self.canvas.draw_idle()

    def _update_graph(self):
        if not self.graph_paused and self.time_buffer:
            self.ax.clear()
            xs = self.time_buffer
            n_total = len(xs)
            # Downsample only if the session has gotten very long, so the
            # plot stays responsive while still showing full time range.
            step = max(1, n_total // 3000)
            xs_ds = xs[::step]
            any_plotted = False
            for i, var in enumerate(self.field_vars):
                if var.get() and i < len(self.field_buffers):
                    ys = self.field_buffers[i][::step]
                    n = min(len(xs_ds), len(ys))
                    if n > 0:
                        self.ax.plot(xs_ds[:n], ys[:n], label=field_label(i))
                        any_plotted = True
            self.ax.set_xlabel("Time (s)")
            self.ax.set_ylabel("Value (decimal)")
            self.ax.autoscale(enable=True, axis="both", tight=False)
            if any_plotted:
                self.ax.legend(loc="upper left", fontsize=8)
            self.canvas.draw_idle()
        self.root.after(GRAPH_REFRESH_MS, self._update_graph)

    # ---------------------------------------------------------------- save
    def _clear_log(self):
        self.raw_rows.clear()
        self.log_text.delete("1.0", "end")
        self.count_var.set("Received: 0")
        self._clear_graph()

    def _save_txt(self):
        if not self.raw_rows:
            messagebox.showinfo(APP_TITLE, "No data to save yet.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                             filetypes=[("Text file", "*.txt")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            for ts, line, _ in self.raw_rows:
                f.write(f"{ts.isoformat()}\t{line}\n")
        messagebox.showinfo(APP_TITLE, f"Saved {len(self.raw_rows)} rows to:\n{path}")

    def _save_xlsx(self):
        if not self.raw_rows:
            messagebox.showinfo(APP_TITLE, "No data to save yet.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel file", "*.xlsx")])
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Raw Log"

        max_fields = max(len(fields) for _, _, fields in self.raw_rows)
        header = ["Timestamp", "Raw Line"] + [field_label(i) for i in range(max_fields)]
        ws.append(header)

        for ts, line, fields in self.raw_rows:
            row = [ts.isoformat(), line] + fields + [""] * (max_fields - len(fields))
            ws.append(row)

        # Reasonable column widths
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 40
        for i in range(max_fields):
            ws.column_dimensions[get_column_letter(3 + i)].width = 10

        wb.save(path)
        messagebox.showinfo(APP_TITLE, f"Saved {len(self.raw_rows)} rows to:\n{path}")

    def _toggle_autosave(self):
        if self.autosave_var.get():
            path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Text file", "*.txt")],
                                                 title="Choose auto-save file")
            if not path:
                self.autosave_var.set(False)
                return
            self.autosave_path = path
            self.autosave_file = open(path, "a", encoding="utf-8")
        else:
            self._toggle_autosave_off()

    def _toggle_autosave_off(self):
        if self.autosave_file:
            try:
                self.autosave_file.close()
            except Exception:
                pass
        self.autosave_file = None
        self.autosave_var.set(False)

    # ---------------------------------------------------------------- exit
    def _on_close(self):
        self._toggle_autosave_off()
        self.running = False
        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    app = SerialMonitorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
