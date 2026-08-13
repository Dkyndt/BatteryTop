# BatteryTop_v10_3_1.py
# Consolidated BatteryTop V7.1
# Based on BatteryTop V6.3 / V7.0
# Features:
# - Moving Average Engine: Avg power 1m / 5m
# - Stable ETA using 5m average when available
# - Battery Health dashboard
# - Two-column dashboard layout
# - Rich Layout grid engine with aligned panels
# - Full-width battery charge bar
# - Equal-height top grid panels
# - Power spike detection
# - Session Analytics panel
# - Right-column sub-layout for Battery Health / Session Analytics / Power Monitor
# - Charge session logging
# - Automatic CSV logging
# - Rebalanced dashboard layout
# - Runtime Prediction engine
# - Power spike history logging
# - Integrated power spike history into Power Monitor
# - Dynamic charge bar colors below battery thresholds
# - Battery delta 1m/5m/10m in power history panels
# - Fixed compact battery delta rows in history panels
# - Alerts & Events panel with battery drop detection
# - Uniform layout alignment and compact history panels
# - Improved Alerts & Events with battery drop warning details
# - One-sample-per-second history guard
# - Corrected percent-based session logging
# - Single full-width mode + percent timeline
# - Compact one-line timeline legend
# - Vertical battery gauge
# - Charge and Discharge histories side-by-side
# Requirements: pip install rich psutil pywin32

import time
import csv
from collections import deque
from datetime import datetime
from pathlib import Path

import psutil
from rich.console import Console, Group
from rich.live import Live
from rich.layout import Layout
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, BarColumn, TextColumn
from rich.text import Text

try:
    import win32com.client
except ImportError:
    raise SystemExit("Install dependencies first: pip install rich psutil pywin32")


console = Console()

VIEW_MODE = "AUTO"

def get_view_mode(console):
    width = console.size.width
    height = console.size.height
    if width < 220 or height < 60:
        return "COMPACT"
    if width < 300:
        return "NORMAL"
    return "FULL"

def timeline_samples_for_mode(mode):
    return 60 if mode=="COMPACT" else 120 if mode=="NORMAL" else 240


LABEL_WIDTH = 18
VALUE_COLUMN_MIN_WIDTH = 12

def make_stats_table(label_width=LABEL_WIDTH, value_justify="right"):
    table = Table(show_header=False, box=None, expand=True)
    table.add_column(width=label_width, style="cyan", no_wrap=True)
    table.add_column(justify=value_justify, overflow="fold")
    return table

REFRESH_SECONDS = 1
HISTORY_LENGTH = 10000
MAX_REASONABLE_POWER_W = 300.0

CSV_LOG_DIR = Path(__file__).resolve().parent / "BatteryLogs"
CSV_LOG_DIR.mkdir(exist_ok=True)
LAST_CSV_LOG_SECOND = None
LAST_HISTORY_SAMPLE_SECOND = None

POWER_SPIKE_FACTOR_THRESHOLD = 2.5
POWER_SPIKE_MIN_INTERVAL_SECONDS = 30
power_spike_history = deque(maxlen=100)
last_power_spike_logged_at = None

BATTERY_DROP_THRESHOLD_PERCENT = 3.0
BATTERY_DROP_WINDOW_SECONDS = 60
BATTERY_DROP_MIN_INTERVAL_SECONDS = 120
BATTERY_DROP_CRITICAL_PERCENT = 10.0
battery_drop_history = deque(maxlen=100)
last_battery_drop_logged_at = None

charge_history = deque(maxlen=HISTORY_LENGTH)
discharge_history = deque(maxlen=HISTORY_LENGTH)
mode_percent_history = deque(maxlen=HISTORY_LENGTH)
power_history_1m = deque(maxlen=60)
power_history_5m = deque(maxlen=300)

last_valid_charge_w = None
last_valid_discharge_w = None
current_mode = None

# Corrected session tracking based on battery percentage deltas.
fixed_completed_sessions = deque(maxlen=200)
fixed_session_mode = None
fixed_session_started_at = None
fixed_session_start_percent = None
fixed_session_powers = []

# Completed charge/discharge session records.
# This is in-memory for the current BatteryTop run.
completed_sessions = deque(maxlen=100)
current_logged_session_mode = None
current_logged_session_started_at = None
current_logged_session_start_percent = None
current_logged_session_powers = []

charge_session_started_at = None
discharge_session_started_at = None
charge_session_start_percent = None
discharge_session_start_percent = None


# -----------------------------
# Helpers
# -----------------------------

def wmi_first(namespace, class_name):
    try:
        path = "winmgmts:{impersonationLevel=impersonate}!\\\\.\\" + namespace
        wmi = win32com.client.GetObject(path)
        for item in wmi.ExecQuery(f"SELECT * FROM {class_name}"):
            return item
    except Exception:
        return None
    return None


def clean_power_mw(raw_value):
    if raw_value is None:
        return None

    try:
        value = float(raw_value)
    except Exception:
        return None

    # Dell / Windows may briefly report sentinel values during plug/unplug transitions.
    if value in (-2147483648, 2147483647):
        return None

    power_w = value / 1000.0

    if power_w < 0 or power_w > MAX_REASONABLE_POWER_W:
        return None

    return power_w


def safe_wh(raw_mwh):
    try:
        if raw_mwh is None:
            return None
        value = float(raw_mwh) / 1000.0
        if value < 0 or value > 1000:
            return None
        return value
    except Exception:
        return None


def safe_voltage(raw_mv):
    try:
        if raw_mv is None:
            return None
        value = float(raw_mv) / 1000.0
        if value <= 0 or value > 30:
            return None
        return value
    except Exception:
        return None


def fmt(value, suffix="", digits=2):
    if value is None:
        return "N/A"
    return f"{value:.{digits}f} {suffix}".rstrip()


def format_elapsed(start_time):
    if start_time is None:
        return "N/A"

    seconds = int((datetime.now() - start_time).total_seconds())
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60

    if hours > 0:
        return f"{hours}h {minutes}m {secs}s"
    if minutes > 0:
        return f"{minutes}m {secs}s"
    return f"{secs}s"


def format_eta(hours):
    if hours is None:
        return "N/A"

    try:
        value = float(hours)
    except Exception:
        return "N/A"

    if value < 0:
        return "N/A"

    total_minutes = int(value * 60)
    h = total_minutes // 60
    m = total_minutes % 60

    return f"{h}h {m}m"


def moving_average(values):
    valid = [v for v in values if v is not None]
    if not valid:
        return None
    return sum(valid) / len(valid)


def moving_average_with_current(values, current):
    valid = [v for v in values if v is not None]
    if current is not None:
        valid.append(current)
    if not valid:
        return None
    return sum(valid) / len(valid)


def calculate_battery_health(design_capacity_wh, full_capacity_wh):
    if design_capacity_wh is None or full_capacity_wh is None:
        return None, None
    if design_capacity_wh <= 0:
        return None, None

    wear_level = ((design_capacity_wh - full_capacity_wh) / design_capacity_wh) * 100.0
    wear_level = max(0.0, wear_level)
    health = max(0.0, 100.0 - wear_level)

    return wear_level, health


def health_status(wear_level):
    if wear_level is None:
        return Text("Unknown", style="white")
    if wear_level < 10:
        return Text("Excellent", style="green")
    if wear_level < 20:
        return Text("Good", style="yellow")
    return Text("Replace Soon", style="red")


def health_border(wear_level):
    if wear_level is None:
        return "cyan"
    if wear_level < 10:
        return "green"
    if wear_level < 20:
        return "yellow"
    return "red"




def detect_power_spike(current_power, avg_power_5m):
    if current_power is None or avg_power_5m is None:
        return False, None

    if avg_power_5m <= 0:
        return False, None

    factor = current_power / avg_power_5m

    return factor >= POWER_SPIKE_FACTOR_THRESHOLD, factor


def build_spike_panel(data):
    current_power = data.get("active_power_w")
    avg_power_5m = data.get("avg_power_5m")
    spike, factor = detect_power_spike(current_power, avg_power_5m)
    recent = list(power_spike_history)[-4:]
    today = spikes_today() if "spikes_today" in globals() else []

    table = make_stats_table()

    table.add_row("Current Power", fmt(current_power, "W", 3))
    table.add_row("Average 5m", fmt(avg_power_5m, "W", 3))
    table.add_row("Factor", f"{factor:.2f}x" if factor is not None else "N/A")

    if spike:
        table.add_row("Status", Text("SPIKE", style="bold red"))
    else:
        table.add_row("Status", Text("Normal", style="green"))

    table.add_row("", "")
    table.add_row("Recent Spikes", "")

    if not recent:
        table.add_row("Recent", "No spikes logged")
    else:
        for item in reversed(recent):
            timestamp = item["timestamp"].strftime("%H:%M:%S")
            power = item.get("current_power_w")
            spike_factor = item.get("factor")
            table.add_row(timestamp, f"{power:.3f} W  {spike_factor:.2f}x")

    if today:
        highest_power = max(item.get("current_power_w") or 0 for item in today)
        largest_factor = max(item.get("factor") or 0 for item in today)
        table.add_row("Today Count", str(len(today)))
        table.add_row("Highest Power", f"{highest_power:.3f} W")
        table.add_row("Largest Factor", f"{largest_factor:.2f}x")
    else:
        table.add_row("Today Count", "0")

    return Panel(
        table,
        title="Power Monitor",
        border_style="red" if spike else "green",
    )


def mode_label(mode):
    if mode == "charge":
        return "Charging"
    if mode == "discharge":
        return "Discharging"
    if mode == "idle":
        return "Idle"
    return "Unknown"


def mode_color(mode):
    if mode == "charge":
        return "green"
    if mode == "discharge":
        return "yellow"
    if mode == "idle":
        return "dim"
    return "white"


def mode_char(mode):
    # V6 style: one full block character per displayed sample.
    # Color shows the battery mode.
    if mode == "charge":
        return "█"
    if mode == "discharge":
        return "█"
    if mode == "idle":
        return "░"
    return "·"


def percent_block(percent):
    # One terminal character per displayed sample.
    # Height shows battery percentage.
    blocks = "▁▂▃▄▅▆▇█"
    if percent is None:
        return "?"
    value = max(0.0, min(float(percent), 100.0))
    index = int(round((value / 100.0) * (len(blocks) - 1)))
    return blocks[index]


def downsample_points(points, width):
    # Compress all stored history into the visible width.
    if width <= 0:
        return []

    if len(points) <= width:
        return list(points)

    step = len(points) / width
    sampled = []

    for i in range(width):
        idx = int(i * step)
        sampled.append(points[idx])

    return sampled


# -----------------------------
# Data collection
# -----------------------------

def get_data():
    status = wmi_first("root\\wmi", "BatteryStatus")
    full = wmi_first("root\\wmi", "BatteryFullChargedCapacity")
    cycles = wmi_first("root\\wmi", "BatteryCycleCount")
    info = wmi_first("root\\cimv2", "Win32_Battery")
    static = wmi_first("root\\wmi", "BatteryStaticData")

    ps_battery = psutil.sensors_battery()

    charging = bool(getattr(status, "Charging", False)) if status else False
    discharging = bool(getattr(status, "Discharging", False)) if status else False
    power_online = bool(getattr(status, "PowerOnline", False)) if status else False

    charge_power_w = clean_power_mw(getattr(status, "ChargeRate", None)) if status else None
    discharge_power_w = clean_power_mw(getattr(status, "DischargeRate", None)) if status else None
    voltage_v = safe_voltage(getattr(status, "Voltage", None)) if status else None
    remaining_wh = safe_wh(getattr(status, "RemainingCapacity", None)) if status else None
    full_capacity_wh = safe_wh(getattr(full, "FullChargedCapacity", None)) if full else None

    # Design capacity is inconsistent across Dell / Windows WMI providers.
    # Try root\\wmi BatteryStaticData first, then Win32_Battery fallback.
    design_capacity_wh = None
    if static is not None:
        design_capacity_wh = safe_wh(getattr(static, "DesignedCapacity", None))
        if design_capacity_wh is None:
            design_capacity_wh = safe_wh(getattr(static, "DesignCapacity", None))
    if design_capacity_wh is None and info is not None:
        design_capacity_wh = safe_wh(getattr(info, "DesignCapacity", None))

    percent = ps_battery.percent if ps_battery else None
    if remaining_wh is not None and full_capacity_wh:
        percent = (remaining_wh / full_capacity_wh) * 100.0

    if percent is not None:
        percent = max(0.0, min(percent, 100.0))

    cycle_count = getattr(cycles, "CycleCount", None) if cycles else None
    battery_name = getattr(info, "Name", "DELL Battery") if info else "DELL Battery"

    charge_active = charging or (power_online and charge_power_w is not None and charge_power_w > 0)
    discharge_active = discharging and not charge_active

    if charge_active:
        mode = "charge"
        active_power_w = charge_power_w
    elif discharge_active:
        mode = "discharge"
        active_power_w = discharge_power_w
    else:
        mode = "idle"
        active_power_w = charge_power_w if charge_power_w is not None else discharge_power_w

    current_a = None
    if voltage_v and active_power_w is not None:
        current_a = active_power_w / voltage_v

    avg_power_1m = moving_average_with_current(power_history_1m, active_power_w)
    avg_power_5m = moving_average_with_current(power_history_5m, active_power_w)
    eta_power_w = avg_power_5m if avg_power_5m and avg_power_5m > 0 else active_power_w

    time_to_full_h = None
    time_to_empty_h = None

    if charge_active and eta_power_w is not None and eta_power_w > 0 and remaining_wh is not None and full_capacity_wh is not None:
        energy_needed_wh = full_capacity_wh - remaining_wh
        if energy_needed_wh > 0:
            time_to_full_h = energy_needed_wh / eta_power_w
        else:
            time_to_full_h = 0

    if discharge_active and eta_power_w is not None and eta_power_w > 0 and remaining_wh is not None:
        time_to_empty_h = remaining_wh / eta_power_w

    wear_level, health = calculate_battery_health(design_capacity_wh, full_capacity_wh)

    return {
        "battery_name": battery_name,
        "percent": percent,
        "charging": charging,
        "discharging": discharging,
        "power_online": power_online,
        "mode": mode,
        "charge_active": charge_active,
        "discharge_active": discharge_active,
        "voltage_v": voltage_v,
        "charge_power_w": charge_power_w,
        "discharge_power_w": discharge_power_w,
        "active_power_w": active_power_w,
        "avg_power_1m": avg_power_1m,
        "avg_power_5m": avg_power_5m,
        "current_a": current_a,
        "remaining_wh": remaining_wh,
        "full_capacity_wh": full_capacity_wh,
        "design_capacity_wh": design_capacity_wh,
        "wear_level": wear_level,
        "health": health,
        "cycle_count": cycle_count,
        "time_to_full_h": time_to_full_h,
        "time_to_empty_h": time_to_empty_h,
    }


# -----------------------------
# History + session handling
# -----------------------------

def update_histories_and_sessions(data):
    global last_valid_charge_w, last_valid_discharge_w
    global current_mode
    global charge_session_started_at, discharge_session_started_at
    global charge_session_start_percent, discharge_session_start_percent
    global LAST_HISTORY_SAMPLE_SECOND

    now = datetime.now()
    sample_second = int(now.timestamp())

    if LAST_HISTORY_SAMPLE_SECOND == sample_second:
        return

    LAST_HISTORY_SAMPLE_SECOND = sample_second

    mode = data["mode"]
    percent = data["percent"]

    update_csv_logging(data, now)
    fixed_update_session_log(data, now)

    if mode != current_mode:
        current_mode = mode
        if mode == "charge":
            charge_session_started_at = now
            charge_session_start_percent = percent
        elif mode == "discharge":
            discharge_session_started_at = now
            discharge_session_start_percent = percent

    charge = data["charge_power_w"]
    discharge = data["discharge_power_w"]
    active_power = data["active_power_w"]

    if active_power is not None:
        power_history_1m.append(active_power)
        power_history_5m.append(active_power)

    if charge is not None:
        last_valid_charge_w = charge

    if discharge is not None:
        last_valid_discharge_w = discharge

    if mode == "charge":
        if charge is not None:
            charge_history.append(charge)
        elif last_valid_charge_w is not None:
            charge_history.append(last_valid_charge_w)
    elif mode == "discharge":
        if discharge is not None:
            discharge_history.append(discharge)
        elif last_valid_discharge_w is not None:
            discharge_history.append(last_valid_discharge_w)

    if percent is not None:
        mode_percent_history.append({
            "time": now,
            "mode": mode,
            "percent": percent,
        })


# -----------------------------
# Charts
# -----------------------------

def sparkline(values, width):
    blocks = "▁▂▃▄▅▆▇█"
    clean = [v for v in values if v is not None]

    if not clean:
        return ""

    data = clean[-width:]

    if len(data) == 1:
        return blocks[3]

    lowest = min(data)
    highest = max(data)

    if highest == lowest:
        return blocks[3] * len(data)

    chars = []
    for value in data:
        index = int((value - lowest) / (highest - lowest) * (len(blocks) - 1))
        chars.append(blocks[index])

    return "".join(chars)





# -----------------------------
# CSV Logging
# -----------------------------

def csv_log_path(now):
    return CSV_LOG_DIR / f"BatteryTop_{now:%Y-%m-%d}.csv"


def csv_row(data, now):
    return {
        "timestamp": now.isoformat(timespec="seconds"),
        "mode": data.get("mode"),
        "battery_percent": data.get("percent"),
        "voltage_v": data.get("voltage_v"),
        "charge_power_w": data.get("charge_power_w"),
        "discharge_power_w": data.get("discharge_power_w"),
        "active_power_w": data.get("active_power_w"),
        "avg_power_1m_w": data.get("avg_power_1m"),
        "avg_power_5m_w": data.get("avg_power_5m"),
        "current_a": data.get("current_a"),
        "remaining_wh": data.get("remaining_wh"),
        "full_capacity_wh": data.get("full_capacity_wh"),
        "design_capacity_wh": data.get("design_capacity_wh"),
        "wear_level_percent": data.get("wear_level"),
        "battery_health_percent": data.get("health"),
        "cycle_count": data.get("cycle_count"),
        "time_to_full_h": data.get("time_to_full_h"),
        "time_to_empty_h": data.get("time_to_empty_h"),
    }


def write_csv_log(data, now):
    path = csv_log_path(now)
    row = csv_row(data, now)
    write_header = not path.exists()

    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(row.keys()))
        if write_header:
            writer.writeheader()
        writer.writerow(row)

    return path


def update_csv_logging(data, now):
    global LAST_CSV_LOG_SECOND

    current_second = int(now.timestamp())
    if LAST_CSV_LOG_SECOND == current_second:
        return None

    LAST_CSV_LOG_SECOND = current_second
    return write_csv_log(data, now)


def build_csv_logging_panel():
    now = datetime.now()
    path = csv_log_path(now)

    table = make_stats_table()
    table.add_row("Status", Text("Enabled", style="green"))
    table.add_row("Folder", str(CSV_LOG_DIR))
    table.add_row("File", path.name)

    return Panel(table, title="CSV Logging", border_style="cyan")


# -----------------------------
# Corrected Percent-Based Session Logging
# -----------------------------

def fixed_finalize_session(end_time, end_percent):
    global fixed_session_mode, fixed_session_started_at, fixed_session_start_percent, fixed_session_powers

    if fixed_session_mode not in ("charge", "discharge"):
        return
    if fixed_session_started_at is None:
        return
    if fixed_session_start_percent is None or end_percent is None:
        return

    values = [value for value in fixed_session_powers if value is not None]
    avg_power = sum(values) / len(values) if values else None
    peak_power = max(values) if values else None
    min_power = min(values) if values else None
    delta_percent = end_percent - fixed_session_start_percent

    if abs(delta_percent) < 0.001 and len(values) < 3:
        return

    fixed_completed_sessions.append({
        "mode": fixed_session_mode,
        "started_at": fixed_session_started_at,
        "ended_at": end_time,
        "start_percent": fixed_session_start_percent,
        "end_percent": end_percent,
        "delta_percent": delta_percent,
        "avg_power": avg_power,
        "peak_power": peak_power,
        "min_power": min_power,
        "samples": len(values),
    })


def fixed_update_session_log(data, now):
    global fixed_session_mode, fixed_session_started_at, fixed_session_start_percent, fixed_session_powers

    mode = data.get("mode")
    percent = data.get("percent")
    active_power = data.get("active_power_w")

    if mode not in ("charge", "discharge"):
        if fixed_session_mode in ("charge", "discharge"):
            fixed_finalize_session(now, percent)
        fixed_session_mode = mode
        fixed_session_started_at = None
        fixed_session_start_percent = None
        fixed_session_powers = []
        return

    if fixed_session_mode != mode:
        fixed_finalize_session(now, percent)
        fixed_session_mode = mode
        fixed_session_started_at = now
        fixed_session_start_percent = percent
        fixed_session_powers = []

    if active_power is not None:
        fixed_session_powers.append(active_power)


def fixed_current_session(data):
    if fixed_session_mode not in ("charge", "discharge"):
        return None

    percent = data.get("percent")
    values = [value for value in fixed_session_powers if value is not None]
    avg_power = sum(values) / len(values) if values else None
    peak_power = max(values) if values else None
    min_power = min(values) if values else None
    delta_percent = None

    if percent is not None and fixed_session_start_percent is not None:
        delta_percent = percent - fixed_session_start_percent

    return {
        "mode": fixed_session_mode,
        "started_at": fixed_session_started_at,
        "start_percent": fixed_session_start_percent,
        "end_percent": percent,
        "delta_percent": delta_percent,
        "avg_power": avg_power,
        "peak_power": peak_power,
        "min_power": min_power,
        "samples": len(values),
    }


def fixed_sessions_today(data):
    today = datetime.now().date()
    sessions = [
        session for session in fixed_completed_sessions
        if session.get("started_at") is not None and session["started_at"].date() == today
    ]

    current = fixed_current_session(data)
    if current is not None and current.get("started_at") is not None and current["started_at"].date() == today:
        sessions.append(current)

    return sessions

# -----------------------------
# Charge / Discharge Session Logging
# -----------------------------

def summarize_power_samples(power_samples):
    values = [v for v in power_samples if v is not None]
    if not values:
        return None, None, None, None

    avg_power = sum(values) / len(values)
    peak_power = max(values)
    min_power = min(values)
    energy_wh = sum(values) * REFRESH_SECONDS / 3600.0

    return avg_power, peak_power, min_power, energy_wh


def finalize_logged_session(end_time, end_percent):
    global current_logged_session_mode
    global current_logged_session_started_at
    global current_logged_session_start_percent
    global current_logged_session_powers

    if current_logged_session_mode not in ("charge", "discharge"):
        return

    if current_logged_session_started_at is None:
        return

    avg_power, peak_power, min_power, energy_wh = summarize_power_samples(
        current_logged_session_powers
    )

    delta_percent = None
    if end_percent is not None and current_logged_session_start_percent is not None:
        delta_percent = end_percent - current_logged_session_start_percent

    completed_sessions.append({
        "mode": current_logged_session_mode,
        "started_at": current_logged_session_started_at,
        "ended_at": end_time,
        "start_percent": current_logged_session_start_percent,
        "end_percent": end_percent,
        "delta_percent": delta_percent,
        "avg_power": avg_power,
        "peak_power": peak_power,
        "min_power": min_power,
        "energy_wh": energy_wh,
        "samples": len(current_logged_session_powers),
    })


def update_charge_session_log(data, now):
    global current_logged_session_mode
    global current_logged_session_started_at
    global current_logged_session_start_percent
    global current_logged_session_powers

    mode = data.get("mode")
    percent = data.get("percent")
    active_power = data.get("active_power_w")

    if mode != current_logged_session_mode:
        finalize_logged_session(now, percent)

        current_logged_session_mode = mode
        current_logged_session_powers = []

        if mode in ("charge", "discharge"):
            current_logged_session_started_at = now
            current_logged_session_start_percent = percent
        else:
            current_logged_session_started_at = None
            current_logged_session_start_percent = None

    if mode in ("charge", "discharge") and active_power is not None:
        current_logged_session_powers.append(active_power)


def current_logged_session_snapshot(data):
    if current_logged_session_mode not in ("charge", "discharge"):
        return None

    avg_power, peak_power, min_power, energy_wh = summarize_power_samples(
        current_logged_session_powers
    )

    current_percent = data.get("percent")
    delta_percent = None
    if current_percent is not None and current_logged_session_start_percent is not None:
        delta_percent = current_percent - current_logged_session_start_percent

    return {
        "mode": current_logged_session_mode,
        "started_at": current_logged_session_started_at,
        "ended_at": None,
        "start_percent": current_logged_session_start_percent,
        "end_percent": current_percent,
        "delta_percent": delta_percent,
        "avg_power": avg_power,
        "peak_power": peak_power,
        "min_power": min_power,
        "energy_wh": energy_wh,
        "samples": len(current_logged_session_powers),
    }


def session_is_today(session):
    started_at = session.get("started_at")
    if started_at is None:
        return False
    return started_at.date() == datetime.now().date()


def aggregate_today_sessions(data):
    sessions = [s for s in completed_sessions if session_is_today(s)]
    current = current_logged_session_snapshot(data)
    if current is not None and session_is_today(current):
        sessions.append(current)

    charge_sessions = [s for s in sessions if s.get("mode") == "charge"]
    discharge_sessions = [s for s in sessions if s.get("mode") == "discharge"]

    charge_energy = sum(s.get("energy_wh") or 0 for s in charge_sessions)
    discharge_energy = sum(s.get("energy_wh") or 0 for s in discharge_sessions)

    charge_delta = sum(s.get("delta_percent") or 0 for s in charge_sessions)
    discharge_delta = sum(s.get("delta_percent") or 0 for s in discharge_sessions)

    return {
        "sessions": sessions,
        "charge_count": len(charge_sessions),
        "discharge_count": len(discharge_sessions),
        "charge_energy": charge_energy,
        "discharge_energy": discharge_energy,
        "charge_delta": charge_delta,
        "discharge_delta": discharge_delta,
        "current": current,
    }


def build_charge_session_log_panel(data):
    sessions = fixed_sessions_today(data)
    charge_sessions = [session for session in sessions if session.get("mode") == "charge"]
    discharge_sessions = [session for session in sessions if session.get("mode") == "discharge"]
    current = fixed_current_session(data)

    charge_gain = sum(max(0.0, session.get("delta_percent") or 0.0) for session in charge_sessions)
    discharge_loss = sum(abs(min(0.0, session.get("delta_percent") or 0.0)) for session in discharge_sessions)

    table = make_stats_table()

    table.add_row("Charge Sessions", str(len(charge_sessions)))
    table.add_row("Discharge Sessions", str(len(discharge_sessions)))
    table.add_row("Charge Gain", fmt(charge_gain, "%", 3))
    table.add_row("Discharge Loss", fmt(discharge_loss, "%", 3))

    if current is not None:
        table.add_row("Current Mode", Text(mode_label(current["mode"]), style=mode_color(current["mode"])))
        table.add_row("Current Since", format_elapsed(current.get("started_at")))
        table.add_row("Current Delta", fmt(current.get("delta_percent"), "%", 3))
        table.add_row("Avg Power", fmt(current.get("avg_power"), "W", 3))
        table.add_row("Peak Power", fmt(current.get("peak_power"), "W", 3))
        table.add_row("Samples", str(current.get("samples") or 0))
    else:
        table.add_row("Current Mode", "Idle")

    return Panel(table, title="Charge Session Log", border_style="cyan")


# -----------------------------
# Battery Trend Analytics
# -----------------------------

def eta_from_power_model(data, power_w):
    if power_w is None or power_w <= 0:
        return None

    remaining_wh = data.get("remaining_wh")
    full_capacity_wh = data.get("full_capacity_wh")
    mode = data.get("mode")

    if mode == "discharge":
        if remaining_wh is None:
            return None
        return remaining_wh / power_w

    if mode == "charge":
        if remaining_wh is None or full_capacity_wh is None:
            return None
        required_wh = max(0.0, full_capacity_wh - remaining_wh)
        return required_wh / power_w

    return None


def percent_rate_per_hour(points, mode_filter=None, max_samples=1800):
    samples = list(points)[-max_samples:]

    if mode_filter is not None:
        samples = [p for p in samples if p.get("mode") == mode_filter]

    if len(samples) < 2:
        return None

    first = samples[0]
    last = samples[-1]

    start_percent = first.get("percent")
    end_percent = last.get("percent")
    start_time = first.get("time")
    end_time = last.get("time")

    if start_percent is None or end_percent is None or start_time is None or end_time is None:
        return None

    seconds = max(1, (end_time - start_time).total_seconds())
    return (end_percent - start_percent) / (seconds / 3600.0)


def eta_from_percent_rate(data, rate_percent_per_hour):
    if rate_percent_per_hour is None or rate_percent_per_hour == 0:
        return None

    percent = data.get("percent")
    mode = data.get("mode")
    if percent is None:
        return None

    if mode == "discharge" and rate_percent_per_hour < 0:
        return percent / abs(rate_percent_per_hour)

    if mode == "charge" and rate_percent_per_hour > 0:
        return (100.0 - percent) / rate_percent_per_hour

    return None


def eta_variance_label(eta_values):
    valid = [v for v in eta_values if v is not None]
    if len(valid) < 2:
        return Text("Unknown", style="white"), None

    spread = max(valid) - min(valid)
    avg = sum(valid) / len(valid)
    relative = spread / avg if avg > 0 else 0

    if relative < 0.10:
        return Text("Low", style="green"), relative
    if relative < 0.25:
        return Text("Medium", style="yellow"), relative
    return Text("High", style="red"), relative


def prediction_confidence(sample_count, eta_values):
    valid = [v for v in eta_values if v is not None]
    if len(valid) < 2 or sample_count < 30:
        return Text("Low", style="red")

    _, relative = eta_variance_label(valid)
    if relative is None:
        return Text("Low", style="red")
    if sample_count >= 300 and relative < 0.10:
        return Text("High", style="green")
    if sample_count >= 60 and relative < 0.25:
        return Text("Medium", style="yellow")
    return Text("Low", style="red")


def weighted_prediction(eta_1m, eta_5m, eta_rate):
    weighted = []
    if eta_1m is not None:
        weighted.append((eta_1m, 0.25))
    if eta_5m is not None:
        weighted.append((eta_5m, 0.50))
    if eta_rate is not None:
        weighted.append((eta_rate, 0.25))

    if not weighted:
        return None

    total_weight = sum(weight for _, weight in weighted)
    return sum(value * weight for value, weight in weighted) / total_weight


def build_battery_trend_panel(data):
    mode = data.get("mode")
    sample_count = len(mode_percent_history)

    eta_current = data.get("time_to_empty_h") if mode == "discharge" else data.get("time_to_full_h") if mode == "charge" else None
    eta_1m = eta_from_power_model(data, data.get("avg_power_1m"))
    eta_5m = eta_from_power_model(data, data.get("avg_power_5m"))

    if mode == "charge":
        rate = percent_rate_per_hour(mode_percent_history, "charge")
    elif mode == "discharge":
        rate = percent_rate_per_hour(mode_percent_history, "discharge")
    else:
        rate = percent_rate_per_hour(mode_percent_history, None)

    eta_rate = eta_from_percent_rate(data, rate)
    predicted_eta = weighted_prediction(eta_1m, eta_5m, eta_rate)
    variance_text, _ = eta_variance_label([eta_1m, eta_5m, eta_rate])
    confidence = prediction_confidence(sample_count, [eta_1m, eta_5m, eta_rate])

    table = make_stats_table()

    table.add_row("Mode", Text(mode_label(mode), style=mode_color(mode)))
    table.add_row("ETA Current|1min|5min|Rate",(
            f"{format_eta(eta_current)} | "
            f"{format_eta(eta_1m)} | "
            f"{format_eta(eta_5m)} | "
            f"{format_eta(eta_rate)}"))
    table.add_row("Predicted", Text(format_eta(predicted_eta), style="bold magenta"))
    table.add_row("Rate", fmt(rate, "%/h", 3))
    table.add_row("Variance", variance_text)
    table.add_row("Confidence", confidence)
    table.add_row("Samples", str(sample_count))

    return Panel(table, title="Runtime Prediction", border_style="magenta")


# -----------------------------
# Session Analytics
# -----------------------------

def build_session_analytics_panel(data):
    current = fixed_current_session(data)
    color = mode_color(data.get("mode"))

    table = make_stats_table()

    if current is None:
        table.add_row("Mode", Text(mode_label(data.get("mode")), style=color))
        table.add_row("Status", "No active charge/discharge session")
        return Panel(table, title="Session Analytics", border_style=color)

    start_percent = current.get("start_percent")
    current_percent = current.get("current_percent")
    delta_percent = current.get("delta_percent")
    started_at = current.get("started_at")

    table.add_row("Mode", Text(mode_label(current["mode"]), style=mode_color(current["mode"])))
    table.add_row("Session", f"{fmt(start_percent, '%', 3)} → {fmt(current_percent, '%', 3)} ({fmt(delta_percent, '%', 3)})")
    table.add_row("Dur/Avg", f"{format_elapsed(started_at)} | {fmt(current.get('avg_power'), 'W', 3)}")
    table.add_row("Pk/Min", f"{fmt(current.get('peak_power'), 'W', 3)} | {fmt(current.get('min_power'), 'W', 3)}")
    table.add_row("Smp", str(current.get("samples") or 0))

    return Panel(
        table,
        title="Session Analytics",
        border_style=mode_color(current["mode"])
    )



# -----------------------------
# Power Spike History
# -----------------------------

def power_spike_log_path(now):
    return CSV_LOG_DIR / f"BatteryTop_PowerSpikes_{now:%Y-%m-%d}.csv"


def write_power_spike_csv(spike):
    path = power_spike_log_path(spike["timestamp"])
    write_header = not path.exists()
    fieldnames = [
        "timestamp",
        "mode",
        "current_power_w",
        "avg_power_5m_w",
        "factor",
        "battery_percent",
    ]

    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        writer.writerow({
            "timestamp": spike["timestamp"].isoformat(timespec="seconds"),
            "mode": spike.get("mode"),
            "current_power_w": spike.get("current_power_w"),
            "avg_power_5m_w": spike.get("avg_power_5m_w"),
            "factor": spike.get("factor"),
            "battery_percent": spike.get("battery_percent"),
        })


def update_power_spike_history(data, now):
    global last_power_spike_logged_at

    current_power = data.get("active_power_w")
    avg_power_5m = data.get("avg_power_5m")
    spike, factor = detect_power_spike(current_power, avg_power_5m)

    if not spike:
        return

    if last_power_spike_logged_at is not None:
        age_seconds = (now - last_power_spike_logged_at).total_seconds()
        if age_seconds < POWER_SPIKE_MIN_INTERVAL_SECONDS:
            return

    record = {
        "timestamp": now,
        "mode": data.get("mode"),
        "current_power_w": current_power,
        "avg_power_5m_w": avg_power_5m,
        "factor": factor,
        "battery_percent": data.get("percent"),
    }

    power_spike_history.append(record)
    last_power_spike_logged_at = now
    write_power_spike_csv(record)


def spikes_today():
    today = datetime.now().date()
    return [
        spike for spike in power_spike_history
        if spike.get("timestamp") is not None and spike["timestamp"].date() == today
    ]


def build_recent_power_spikes_panel():
    recent = list(power_spike_history)[-5:]
    today = spikes_today()

    table = make_stats_table()

    if not recent:
        table.add_row("Recent", "No spikes logged")
    else:
        for spike in reversed(recent):
            timestamp = spike["timestamp"].strftime("%H:%M:%S")
            power = spike.get("current_power_w")
            factor = spike.get("factor")
            table.add_row(timestamp, f"{power:.3f} W  {factor:.2f}x")

    if today:
        highest_power = max(s.get("current_power_w") or 0 for s in today)
        largest_factor = max(s.get("factor") or 0 for s in today)
        table.add_row("Today Count", str(len(today)))
        table.add_row("Highest Power", f"{highest_power:.3f} W")
        table.add_row("Largest Factor", f"{largest_factor:.2f}x")
    else:
        table.add_row("Today Count", "0")

    return Panel(
        table,
        title="Recent Power Spikes",
        border_style="red" if recent else "green",
    )


# -----------------------------
# Battery Thresholds and Delta Analytics
# -----------------------------

def battery_charge_color(data):
    percent = data.get("percent")

    if percent is None:
        return "white"

    if percent <= 5:
        return "bright_red"
    elif percent <= 15:
        return "red"
    elif percent <= 30:
        return "yellow"
    else:
        return "green"


def battery_percent_delta(minutes, mode_filter=None):
    if not mode_percent_history:
        return None

    now = datetime.now()
    cutoff_seconds = minutes * 60
    samples = []

    for point in reversed(mode_percent_history):
        if point.get("time") is None or point.get("percent") is None:
            continue

        age = (now - point["time"]).total_seconds()
        if age > cutoff_seconds:
            break

        if mode_filter is not None and point.get("mode") != mode_filter:
            continue

        samples.append(point)

    if len(samples) < 2:
        return None

    newest = samples[0]
    oldest = samples[-1]
    return newest["percent"] - oldest["percent"]


def fmt_delta_percent(value):
    if value is None:
        return "N/A"
    return f"{value:+.3f} %"


def add_delta_rows(stats, mode_filter):
    delta_1m = fmt_delta_percent(battery_percent_delta(1, mode_filter))
    delta_5m = fmt_delta_percent(battery_percent_delta(5, mode_filter))
    delta_10m = fmt_delta_percent(battery_percent_delta(10, mode_filter))

    def compact(value):
        if value is None:
            return "N/A"

        text = str(value)
        return text.replace(" %", "").replace("%", "")

    stats.add_row("Δ 1/5/10m",(
            f"{compact(delta_1m)} % | "
            f"{compact(delta_5m)} % | "
            f"{compact(delta_10m)} %"),)


# -----------------------------
# Battery Drop Detection + Alerts & Events
# -----------------------------

def battery_drop_log_path(now):
    return CSV_LOG_DIR / f"BatteryTop_BatteryDrops_{now:%Y-%m-%d}.csv"


def write_battery_drop_csv(drop_event):
    path = battery_drop_log_path(drop_event["timestamp"])
    write_header = not path.exists()
    fieldnames = [
        "timestamp",
        "previous_percent",
        "current_percent",
        "drop_percent",
        "window_seconds",
        "mode",
        "active_power_w",
        "avg_power_5m_w",
    ]

    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        writer.writerow({
            "timestamp": drop_event["timestamp"].isoformat(timespec="seconds"),
            "previous_percent": drop_event.get("previous_percent"),
            "current_percent": drop_event.get("current_percent"),
            "drop_percent": drop_event.get("drop_percent"),
            "window_seconds": drop_event.get("window_seconds"),
            "mode": drop_event.get("mode"),
            "active_power_w": drop_event.get("active_power_w"),
            "avg_power_5m_w": drop_event.get("avg_power_5m_w"),
        })


def update_battery_drop_history(data, now):
    global last_battery_drop_logged_at

    current_percent = data.get("percent")
    if current_percent is None or not mode_percent_history:
        return

    # Find the oldest available sample inside the configured detection window.
    reference = None
    for point in reversed(mode_percent_history):
        if point.get("time") is None or point.get("percent") is None:
            continue

        age = (now - point["time"]).total_seconds()
        if age > BATTERY_DROP_WINDOW_SECONDS:
            break

        reference = point

    if reference is None:
        return

    previous_percent = reference.get("percent")
    if previous_percent is None:
        return

    drop_percent = current_percent - previous_percent
    if drop_percent > -BATTERY_DROP_THRESHOLD_PERCENT:
        return

    if last_battery_drop_logged_at is not None:
        age = (now - last_battery_drop_logged_at).total_seconds()
        if age < BATTERY_DROP_MIN_INTERVAL_SECONDS:
            return

    window_seconds = max(1, int((now - reference["time"]).total_seconds()))
    event = {
        "timestamp": now,
        "previous_percent": previous_percent,
        "current_percent": current_percent,
        "drop_percent": drop_percent,
        "window_seconds": window_seconds,
        "mode": data.get("mode"),
        "active_power_w": data.get("active_power_w"),
        "avg_power_5m_w": data.get("avg_power_5m"),
    }

    battery_drop_history.append(event)
    last_battery_drop_logged_at = now
    write_battery_drop_csv(event)


def battery_drops_today():
    today = datetime.now().date()
    return [
        event for event in battery_drop_history
        if event.get("timestamp") is not None and event["timestamp"].date() == today
    ]


def build_alerts_events_panel():
    table = make_stats_table()

    drops_today = battery_drops_today()
    spikes_today_list = spikes_today() if "spikes_today" in globals() else []

    events = []

    for drop in battery_drop_history:
        drop_percent = drop.get("drop_percent") or 0
        severity = "CRITICAL DROP" if abs(drop_percent) >= BATTERY_DROP_CRITICAL_PERCENT else "BATTERY DROP"
        events.append({
            "timestamp": drop["timestamp"],
            "kind": "drop",
            "label": severity,
            "detail": f"{drop_percent:+.3f} %",
            "style": "bold red" if severity == "CRITICAL DROP" else "red",
            "event": drop,
        })

    for spike in power_spike_history:
        events.append({
            "timestamp": spike["timestamp"],
            "kind": "spike",
            "label": "Power Spike",
            "detail": f"{spike.get('factor'):.2f}x",
            "style": "yellow",
            "event": spike,
        })

    events = sorted(events, key=lambda item: item["timestamp"], reverse=True)
    latest = events[0] if events else None

    has_drop = bool(drops_today)
    has_critical = any(abs(event.get("drop_percent") or 0) >= BATTERY_DROP_CRITICAL_PERCENT for event in drops_today)
    has_spike = bool(spikes_today_list)

    if has_critical:
        status_text = Text("CRITICAL DROP", style="bold red")
        border = "bright_red"
    elif has_drop:
        status_text = Text("BATTERY DROP", style="red")
        border = "red"
    elif has_spike:
        status_text = Text("POWER SPIKE", style="yellow")
        border = "yellow"
    else:
        status_text = Text("Healthy", style="green")
        border = "green"

    table.add_row("Status", status_text)
    table.add_row("Drops Today", str(len(drops_today)))
    table.add_row("Spikes Today", str(len(spikes_today_list)))
    table.add_row("Drop Rule", f">= {BATTERY_DROP_THRESHOLD_PERCENT:.1f} % / {BATTERY_DROP_WINDOW_SECONDS}s")

    if latest is None:
        table.add_row("Latest", Text("No alerts", style="green"))
    else:
        table.add_row("Latest", Text(latest["label"], style=latest["style"]))
        table.add_row("Time", latest["timestamp"].strftime("%H:%M:%S"))

        if latest["kind"] == "drop":
            drop = latest["event"]
            previous_percent = drop.get("previous_percent")
            current_percent = drop.get("current_percent")
            drop_percent = drop.get("drop_percent")
            window_seconds = drop.get("window_seconds")

            if previous_percent is not None and current_percent is not None:
                table.add_row("Change", f"{previous_percent:.3f} % -> {current_percent:.3f} %")
            table.add_row("Drop", f"{drop_percent:+.3f} %" if drop_percent is not None else "N/A")
            table.add_row("Window", f"{window_seconds}s" if window_seconds is not None else "N/A")
            table.add_row("Mode", mode_label(drop.get("mode")))
            table.add_row("Power", fmt(drop.get("active_power_w"), "W", 3))

        elif latest["kind"] == "spike":
            spike = latest["event"]
            table.add_row("Factor", f"{spike.get('factor'):.2f}x")
            table.add_row("Power", fmt(spike.get("current_power_w"), "W", 3))

    if drops_today:
        largest_drop = min(event.get("drop_percent") or 0 for event in drops_today)
        table.add_row("Largest Drop", f"{largest_drop:+.3f} %")

    return Panel(table, title="Alerts & Events", border_style=border)

# -----------------------------
# UI rendering
# -----------------------------

def build_battery_panel(data):
    table = make_stats_table()

    table.add_row("Battery", str(data["battery_name"]))
    table.add_row("Status", Text(mode_label(data["mode"]), style=mode_color(data["mode"])))
    table.add_row("Battery level", fmt(data["percent"], "%", 3))
    table.add_row("Voltage", fmt(data["voltage_v"], "V", 3))
    table.add_row("Chg/Dischg",(f"{fmt(data['charge_power_w'], 'W', 3)} | "f"{fmt(data['discharge_power_w'], 'W', 3)}"),)
    table.add_row("Avg Pwr 1m/5m",(f"{fmt(data['avg_power_1m'], 'W', 3)} | "f"{fmt(data['avg_power_5m'], 'W', 3)}"))
    table.add_row("Active current", fmt(data["current_a"], "A", 3))
    table.add_row("Remaining", fmt(data["remaining_wh"], "Wh", 3))
    table.add_row("Full capacity", fmt(data["full_capacity_wh"], "Wh", 3))
    table.add_row("Cycle count", str(data["cycle_count"]) if data["cycle_count"] is not None else "N/A")

    time_to_full = format_eta(data.get("time_to_full_h"))
    time_to_empty = format_eta(data.get("time_to_empty_h"))

    if data["mode"] == "charge":
        table.add_row("Time to full", Text(time_to_full, style="green"))
    else:
        table.add_row("Time to full", time_to_full)

    if data["mode"] == "discharge":
        table.add_row("Time to empty", Text(time_to_empty, style="yellow"))
    else:
        table.add_row("Time to empty", time_to_empty)

    return Panel(table, title="BatteryTop V10.3.1 - Dell")


def build_battery_health_panel(data):
    table = make_stats_table()

    wear = data.get("wear_level")
    health = data.get("health")

    table.add_row("Design Capacity", fmt(data.get("design_capacity_wh"), "Wh", 3))
    table.add_row("Current Capacity", fmt(data.get("full_capacity_wh"), "Wh", 3))
    table.add_row("Remaining", fmt(data.get("remaining_wh"), "Wh", 3))
    table.add_row("Wear Level", fmt(wear, "%", 3))
    table.add_row("Battery Health", fmt(health, "%", 3))
    table.add_row("Cycle Count", str(data.get("cycle_count")) if data.get("cycle_count") is not None else "N/A")
    table.add_row("Status", health_status(wear))

    return Panel(
        table,
        title="Battery Health",
        border_style=health_border(wear),
    )


def build_vertical_battery_panel(data):
    percent = data.get("percent")
    color = battery_charge_color(data)

    try:
        value = float(percent) if percent is not None else 0.0
    except Exception:
        value = 0.0

    value = max(0.0, min(value, 100.0))

    # Dynamic full-height gauge.
    # The gauge uses the full available terminal height and a wider battery body.
    reserved_rows = 8
    total_rows = max(10, console.size.height - reserved_rows)
    filled_rows = int(round((value / 100.0) * total_rows))

    lines = []

    top_label = Text()
    top_label.append("100%", style="cyan")
    lines.append(top_label)

    # Render from top to bottom. 0% stays at the bottom, 100% at the top.
    # Empty cells render as a battery shell. Filled cells render as a wide solid block.
    for row_index in range(total_rows):
        level_from_bottom = total_rows - row_index
        filled = level_from_bottom <= filled_rows

        line = Text()
        if filled:
            line.append("│█████│", style=color)
        else:
            line.append("│     │", style="dim")
        lines.append(line)

    bottom_label = Text()
    bottom_label.append("  0%", style="cyan")
    lines.append(bottom_label)

    percent_text = Text(f"{value:>6.2f}%", style=f"bold {color}")
    mode_text = Text(mode_label(data.get("mode")), style=mode_color(data.get("mode")))

    return Panel(
        Group(*lines, "", percent_text, mode_text),
        title="Battery",
        border_style=color,
        padding=(0, 1),
    )


def build_main_progress(data):
    color = battery_charge_color(data)

    progress = Progress(
        TextColumn("[cyan]Battery[/cyan]"),
        BarColumn(
            bar_width=None,
            complete_style=color,
            finished_style=color,
        ),
        TextColumn("[bold]{task.percentage:>6.3f}%[/bold]"),
        expand=True,
    )

    progress.add_task(
        "battery",
        total=100,
        completed=max(0, min(data["percent"] or 0, 100)),
    )

    return Panel(
        progress,
        title="Battery Charge",
        border_style=color,
        padding=(0, 1),
    )


def build_history_panel(history_values, title, color, active, session_started_at, session_start_percent, current_percent):
    valid = [v for v in history_values if v is not None]

    if not valid:
        return Panel("No power history yet", title=title, border_style=color)

    graph_width = max(8, (console.size.width // 2) - 14)
    graph = sparkline(valid, graph_width)

    current = valid[-1]
    minimum = min(valid)
    maximum = max(valid)
    average = sum(valid) / len(valid)

    session_delta_percent = None
    if current_percent is not None and session_start_percent is not None:
        session_delta_percent = current_percent - session_start_percent

    if "Charge" in title:
        delta_mode_filter = "charge"
    elif "Discharge" in title:
        delta_mode_filter = "discharge"
    else:
        delta_mode_filter = None

    graph_text = Text(graph, style=color, no_wrap=True)

    stats = make_stats_table()
    value_text = Text()
    value_text.append(f"{current:.3f} W", style=f"bold {color}")
    value_text.append(f" | {minimum:.3f} W | {maximum:.3f} W | {average:.3f} W")
    stats.add_row("Cur/Min/Max/Avg", value_text)
    add_delta_rows(stats, delta_mode_filter)
    session_text = (f"{fmt(session_start_percent, '%', 3)}"
        f" → "
        f"{fmt(current_percent, '%', 3)}"
        f" ({fmt(session_delta_percent, '%', 3)})"
    )
    stats.add_row("Session", session_text,)
    stats.add_row("Duration",format_elapsed(session_started_at),)

    status_icon = "●" if active else "○"
    panel_title = f"{title} {status_icon}"
    content = Group(graph_text, stats)

    return Panel(content, title=panel_title, border_style=color)


def build_mode_percent_timeline_panel(data):
    valid = list(mode_percent_history)

    if not valid:
        return Panel("No battery timeline yet", title="Battery Mode + Percent Timeline", border_style="cyan")

    graph_width = max(8, console.size.width - 22)
    samples = downsample_points(valid, graph_width)

    mode_line = Text(no_wrap=True)
    percent_line = Text(no_wrap=True)

    for point in samples:
        color = mode_color(point["mode"])
        mode_line.append(mode_char(point["mode"]), style=color)
        percent_line.append(percent_block(point["percent"]), style=color)

    first = valid[0]
    last = valid[-1]
    delta = last["percent"] - first["percent"]

    legend = Text(no_wrap=True)
    legend.append("Legend ", style="cyan")
    legend.append("█ charge ", style="green")
    legend.append("█ discharge ", style="yellow")
    legend.append("░ idle ", style="dim")
    legend.append(f"({len(valid)} samples)", style="white")

    stats = make_stats_table()
    stats.add_row("Session",(
        f"{fmt(first['percent'], '%', 3)}"
        f" → "
        f"{fmt(last['percent'], '%', 3)}"
        f" ({fmt(delta, '%', 3)})"    ),)
    stats.add_row(    "Started",    first["time"].strftime("%H:%M:%S"))
    stats.add_row(    "Duration",    format_elapsed(first["time"]))
    stats.add_row(    "Mode",    Text(mode_label(data["mode"]), style=mode_color(data["mode"])))

    content = Group(
        Text("Mode    ", style="cyan") + mode_line,
        Text("Battery ", style="cyan") + percent_line,
        legend,
        stats,
    )

    return Panel(content, title="Battery Mode + Percent Timeline", border_style="cyan")


def build_dashboard_layout(data):
    layout = Layout(name="root")

    # V10.2 layout:
    # Left: dynamic full-height vertical battery gauge, 0% bottom and 100% top.
    # Right: existing dashboard grid without the old horizontal Battery Charge row.
    layout.split_row(
        Layout(name="battery_gauge", size=14),
        Layout(name="main_dashboard", ratio=1),
    )

    layout["main_dashboard"].split_column(
        Layout(name="row_status", ratio=5, minimum_size=12),
        Layout(name="row_health", ratio=5, minimum_size=12),
        Layout(name="row_timeline", ratio=2, minimum_size=7),
        Layout(name="row_history", ratio=2, minimum_size=6),
    )

    layout["row_status"].split_row(
        Layout(name="battery_status", ratio=1),
        Layout(name="right_status", ratio=1),
    )

    layout["right_status"].split_column(
        Layout(name="session_analytics", ratio=1, minimum_size=6),
        Layout(name="charge_session_log", ratio=1, minimum_size=6),
    )

    layout["row_health"].split_row(
        Layout(name="left_health", ratio=1),
        Layout(name="right_health", ratio=1),
    )

    layout["left_health"].split_column(
        Layout(name="battery_health", ratio=1),
        Layout(name="battery_trend", ratio=1),
    )

    layout["right_health"].split_column(
        Layout(name="power_monitor", ratio=2, minimum_size=8),
        Layout(name="alerts_events", ratio=1, minimum_size=4),
    )

    layout["row_history"].split_row(
        Layout(name="charge_history", ratio=1),
        Layout(name="discharge_history", ratio=1),
    )

    layout["battery_gauge"].update(build_vertical_battery_panel(data))

    layout["battery_status"].update(build_battery_panel(data))
    layout["session_analytics"].update(build_session_analytics_panel(data))
    layout["charge_session_log"].update(build_charge_session_log_panel(data))

    layout["battery_health"].update(build_battery_health_panel(data))
    layout["battery_trend"].update(build_battery_trend_panel(data))
    layout["power_monitor"].update(build_spike_panel(data))
    layout["alerts_events"].update(build_alerts_events_panel())

    layout["row_timeline"].update(build_mode_percent_timeline_panel(data))

    layout["charge_history"].update(
        build_history_panel(
            charge_history,
            "Charge Power History",
            "green",
            data["mode"] == "charge",
            charge_session_started_at,
            charge_session_start_percent,
            data["percent"],
        )
    )

    layout["discharge_history"].update(
        build_history_panel(
            discharge_history,
            "Discharge Power History",
            "yellow",
            data["mode"] == "discharge",
            discharge_session_started_at,
            discharge_session_start_percent,
            data["percent"],
        )
    )

    return layout


def ui():
    data = get_data()
    update_histories_and_sessions(data)
    return build_dashboard_layout(data)

def main():
    try:
        with Live(ui(), refresh_per_second=2) as live:
            while True:
                live.update(ui())
                time.sleep(REFRESH_SECONDS)
    except KeyboardInterrupt:
        console.print("\nBatteryTop closed.")


if __name__ == "__main__":
    main()
