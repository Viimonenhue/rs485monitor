"""
RS-485 / USB Serial Monitor
---------------------------
A lightweight Windows desktop tool to connect to an RS-485-to-USB adapter,
show incoming hex-formatted lines live in a table, plot selected fields on
a live graph, and save the raw log to .txt or .xlsx.

Build into a standalone .exe with PyInstaller (see README.txt).
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
from datetime import datetime

import serial
import serial.tools.list_ports

import openpyxl
from openpyxl.utils import get_column_letter

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

APP_TITLE = "RS-485 Serial Monitor"
APP_VERSION = "1.0"   # bump this each time I hand you a new file
MAX_TABLE_ROWS = 4000            # cap what's ON SCREEN; saved data is never trimmed
GRAPH_REFRESH_MS = 300           # how often the graph redraws
QUEUE_POLL_MS = 50               # how often we drain incoming serial data
SAMPLE_INTERVAL_S = 0.02         # one line = 0.02s (device's fixed cycle, ~50/sec)

# Named meaning of each semicolon-separated field, in order (A..L in your
# Excel breakdown).
FIELD_NAMES = [
    "Encoder 1",         # A
    "Encoder 2",          # B
    "Effekt kW M1",       # C
    "Effekt kW M2",        # D
    "Strom A M1",         # E
    "Strom A M2",          # F
    "C211 M1",            # G
    "C211 M2",             # H
    "C212 M1",            # I
    "C212 M2",             # J
    "Port B",              # K
    "Felkod",               # L
]

# Display-batch options: how many lines accumulate before the table repaints.
# Real data rate is ~50 lines/sec; these keep the UI from redrawing 50x/sec.
BATCH_OPTIONS = [
    ("0.2s  (flush every 10 lines)", 10),
    ("15s  (flush every 50 lines)", 50),
    ("30s  (flush every 100 lines)", 100),
    ("1 min  (flush every 250 lines)", 250),
    ("5 min+  (flush every 1000 lines)", 1000),
]


def field_label(index):
    """Human-readable label for field position `index` (0-based)."""
    if index < len(FIELD_NAMES):
        return FIELD_NAMES[index]
    return f"Field {index + 1}"


def field_scale_info(name):
    """Returns (scale_factor, unit, cap) for a given field name, used only
    for the GRAPH display. Raw hex values in the table/save are untouched.

    - Effekt kW fields: raw value is in units of 0.01 kW -> divide by 100,
      capped at 1.99 kW (spikes above that are almost certainly noise/glitch
      reads, not real readings).
    - Strom A (current) fields: raw value is in units of 0.1 A -> divide by
      10, capped at 20 A.
    - Everything else: shown as-is, no unit, no cap.
    """
    if "kW" in name:
        return 0.01, "kW", 1.99
    if name.startswith("Strom"):
        return 0.1, "A", 20.0
    return 1.0, None, None


class SerialMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.root.geometry("1250x720")

        self.ser = None
        self.read_thread = None
        self.running = False
        self.data_queue = queue.Queue()

        # Full history of received lines: (timestamp, raw_line, [fields]).
        # This is NEVER trimmed - it's the source of truth for Save.
        self.raw_rows = []

        # Graph state - keeps full history too; redraw is downsampled for
        # performance, not the underlying data.
        self.field_count = 12
        self.field_vars = []
        self.field_buffers = []
        self.time_buffer = []
        self.sample_index = 0
        self.graph_paused = False

        # Some devices only transmit a field once and omit it (blank) on
        # later lines if it hasn't changed. We carry forward the last known
        # value for DISPLAY/GRAPH purposes only - the saved raw file always
        # stores exactly what was received, blank or not.
        self.last_known = [""] * self.field_count

        # Table batching state
        self.pending_rows = []     # rows waiting to be flushed to the table
        self.batch_size = 10

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
        # ---- Left settings panel ----
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

        ttk.Label(settings, text="Display Batching", font=("Segoe UI", 11, "bold")).grid(
            row=r, column=0, columnspan=2, sticky="w")
        r += 1
        ttk.Label(settings,
                  text="Real data rate is ~50 lines/sec.\nTable repaints once this many\nlines have arrived (nothing is\never discarded).",
                  font=("Segoe UI", 8), foreground="#555").grid(row=r, column=0, columnspan=2, sticky="w", pady=(2, 4))
        r += 1
        self.batch_var = tk.StringVar(value=BATCH_OPTIONS[0][0])
        batch_combo = ttk.Combobox(settings, textvariable=self.batch_var, width=26, state="readonly",
                                    values=[b[0] for b in BATCH_OPTIONS])
        batch_combo.grid(row=r, column=0, columnspan=2, sticky="we")
        batch_combo.bind("<<ComboboxSelected>>", self._on_batch_change)
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
        ttk.Button(settings, text="Clear All", command=self._clear_all).grid(
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

        self._build_log_tab(notebook)
        self._build_graph_tab(notebook)

    # --------------------------------------------------------------- Log tab
    def _build_log_tab(self, notebook):
        log_tab = ttk.Frame(notebook)
        notebook.add(log_tab, text="Live Log")

        top_row = ttk.Frame(log_tab)
        top_row.pack(anchor="w", fill="x", padx=5, pady=(5, 2))
        self.autoscroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top_row, text="Auto-scroll", variable=self.autoscroll_var).pack(side="left")

        # Tighter row height / smaller font via a dedicated style
        style = ttk.Style()
        style.configure("Tight.Treeview", font=("Consolas", 9), rowheight=18)
        style.configure("Tight.Treeview.Heading", font=("Consolas", 9, "bold"))

        table_frame = ttk.Frame(log_tab)
        table_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        self.log_tree = ttk.Treeview(table_frame, show="headings", style="Tight.Treeview")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.log_tree.yview)
        self.log_tree.configure(yscrollcommand=yscroll.set)
        self.log_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.log_tree.tag_configure("fault", foreground="#c0392b", background="#fdecea")

        self._rebuild_table_columns()

    def _rebuild_table_columns(self):
        cols = ["elapsed"] + [f"f{i}" for i in range(self.field_count)]
        self.log_tree.configure(columns=cols)
        self.log_tree.column("elapsed", width=80, anchor="e", stretch=False)
        self.log_tree.heading("elapsed", text="Time (s)")
        for i in range(self.field_count):
            key = f"f{i}"
            self.log_tree.column(key, width=90, anchor="center", stretch=True)
            self.log_tree.heading(key, text=field_label(i))

    # ------------------------------------------------------------- Graph tab
    def _build_graph_tab(self, notebook):
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
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_tab)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)
        self._style_axes()

    def _build_field_checkboxes(self):
        for w in self.field_checks_frame.winfo_children():
            w.destroy()
        self.field_vars = []
        self.field_buffers = []
        for i in range(self.field_count):
            var = tk.BooleanVar(value=(i == 2))  # Effekt kW M1 ticked by default
            cb = ttk.Checkbutton(self.field_checks_frame, text=field_label(i), variable=var,
                                  command=self._update_graph_ylabel_now)
            cb.pack(side="left", padx=(0, 4))
            self.field_vars.append(var)
            self.field_buffers.append([])

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

    def _on_batch_change(self, event=None):
        for label, size in BATCH_OPTIONS:
            if label == self.batch_var.get():
                self.batch_size = size
                return

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
        while not self.data_queue.empty() and drained < 1000:
            line = self.data_queue.get_nowait()
            self._handle_line(line)
            drained += 1
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def _handle_line(self, line):
        fields = [f.strip() for f in line.split(";")]

        if len(fields) != self.field_count and len(fields) > 0:
            self.field_count = len(fields)
            self.last_known = [""] * self.field_count
            self._build_field_checkboxes()
            self._rebuild_table_columns()

        elapsed = round(self.sample_index * SAMPLE_INTERVAL_S, 2)
        ts = datetime.now()

        # `fields` = exactly what was received, blanks and all - this is
        # what gets saved to file, never altered.
        self.raw_rows.append((ts, line, fields))

        # `display_fields` = same, but blank entries are filled in with the
        # last known non-blank value, for the table and graph only.
        display_fields = []
        for i, f in enumerate(fields):
            if f == "" and i < len(self.last_known) and self.last_known[i] != "":
                display_fields.append(self.last_known[i])
            else:
                display_fields.append(f)
                if f != "" and i < len(self.last_known):
                    self.last_known[i] = f

        fault = False
        if display_fields:
            try:
                fault = int(display_fields[-1], 16) != 0
            except ValueError:
                fault = False

        # Queue this row for the next table flush (never dropped - just
        # waiting its turn to be painted).
        self.pending_rows.append((elapsed, display_fields, fault))
        if len(self.pending_rows) >= self.batch_size:
            self._flush_table()

        self.count_var.set(f"Received: {len(self.raw_rows)}")

        # ---- Graph buffers (full history kept, scaled/capped per field) ----
        self.time_buffer.append(elapsed)
        self.sample_index += 1
        for i, f in enumerate(display_fields):
            if i >= len(self.field_buffers):
                break
            try:
                raw_val = int(f, 16)
            except ValueError:
                raw_val = None
            if raw_val is not None:
                scale, _unit, cap = field_scale_info(field_label(i))
                val = raw_val * scale
                if cap is not None and val > cap:
                    val = cap
            else:
                # Blank/unparseable field from the device - store as NaN so
                # matplotlib just draws a gap instead of erroring out.
                val = float("nan")
            self.field_buffers[i].append(val)

        # ---- Auto-save ----
        if self.autosave_file:
            self.autosave_file.write(f"{ts.isoformat()}\t{line}\n")
            self.autosave_file.flush()

    def _flush_table(self):
        if not self.pending_rows:
            return
        for elapsed, fields, fault in self.pending_rows:
            values = [f"{elapsed:.2f}"] + fields
            tags = ("fault",) if fault else ()
            self.log_tree.insert("", "end", values=values, tags=tags)
        self.pending_rows.clear()

        # Trim on-screen rows only - raw_rows (used for Save) is untouched.
        children = self.log_tree.get_children()
        overflow = len(children) - MAX_TABLE_ROWS
        if overflow > 0:
            for item in children[:overflow]:
                self.log_tree.delete(item)

        if self.autoscroll_var.get():
            children = self.log_tree.get_children()
            if children:
                self.log_tree.see(children[-1])

    # --------------------------------------------------------------- graph
    def _toggle_pause(self):
        self.graph_paused = not self.graph_paused
        self.pause_btn.configure(text="Resume Graph" if self.graph_paused else "Pause Graph")

    def _clear_graph(self):
        self.time_buffer.clear()
        for buf in self.field_buffers:
            buf.clear()
        self.ax.clear()
        self._style_axes()
        self.canvas.draw_idle()

    def _checked_field_names(self):
        return [field_label(i) for i, var in enumerate(self.field_vars) if var.get()]

    def _checked_units(self):
        units = set()
        for name in self._checked_field_names():
            _, unit, _ = field_scale_info(name)
            units.add(unit)
        return units

    def _ylabel_for_checked(self):
        units = self._checked_units()
        if len(units) == 1:
            unit = next(iter(units))
            if unit == "kW":
                return "Power (kW)"
            if unit == "A":
                return "Current (A)"
        return "Value (decimal)"

    def _style_axes(self):
        """Common axis styling applied after every ax.clear()."""
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel(self._ylabel_for_checked())
        self.ax.grid(True, linewidth=0.3, alpha=0.6)

    def _update_graph_ylabel_now(self):
        # Called immediately on checkbox click so the label feels responsive,
        # even though the plot itself redraws on the timer.
        self.ax.set_ylabel(self._ylabel_for_checked())
        self.canvas.draw_idle()

    def _update_graph(self):
        if not self.graph_paused and self.time_buffer:
            self.ax.clear()
            xs = self.time_buffer
            n_total = len(xs)
            # Downsample only for very long sessions so redraw stays fast;
            # the underlying data itself is never discarded.
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
            self._style_axes()
            if any_plotted:
                self.ax.legend(loc="upper left", fontsize=8)
            self.canvas.draw_idle()
        self.root.after(GRAPH_REFRESH_MS, self._update_graph)

    # ---------------------------------------------------------------- save
    def _clear_all(self):
        self.raw_rows.clear()
        self.pending_rows.clear()
        for item in self.log_tree.get_children():
            self.log_tree.delete(item)
        self.sample_index = 0
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
