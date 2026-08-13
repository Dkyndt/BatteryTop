# BatteryTop

Real-time battery analytics dashboard for Windows laptops.

BatteryTop provides advanced battery monitoring, runtime prediction, battery health analytics, power spike detection and battery drop warnings in a compact terminal dashboard.

<p align="center">
  docs/images/dashboard.png
</p>

---

## Features

### Battery Monitoring

Monitor battery performance in real time:

- Battery percentage
- Voltage
- Current
- Charge power
- Discharge power
- Remaining capacity
- Estimated runtime

---

### Battery Health

Track battery condition:

- Design Capacity
- Current Capacity
- Remaining Capacity
- Wear Level

---

### Runtime Prediction

BatteryTop calculates predicted runtime using multiple models:

- Current Windows estimate
- 1-minute power model
- 5-minute power model

This provides more stable runtime estimates than the standard Windows battery indicator.

---

### Session Analytics

Track charging and discharging sessions:

- Session duration
- Start percentage
- Current percentage
- Session delta
- Average power
- Peak power

---

### Charge Session Log

Daily charging statistics:

- Charge sessions
- Discharge sessions
- Total charge gain
- Total discharge loss

---

### Power Monitor

Monitor power consumption:

- Current power draw
- 5-minute average
- Spike factor
- Recent power spikes
- Daily statistics

---

### Power Spike Detection

Automatically detects unusual power consumption peaks.

Examples:

- Visual Studio builds
- Teams meetings
- OneDrive synchronization
- Windows updates

---

### Battery Drop Detection

Detects abnormal battery percentage drops.

Example:

```text
15.4% → 5.2%

Drop: -10.2%
```

BatteryTop logs and reports these events automatically.

---

### Alerts & Events

Centralized warning panel showing:

- Battery drops
- Critical battery drops
- Power spikes
- Daily statistics
- Latest event details

Status indicators:

```text
Healthy
Power Spike
Battery Drop
Critical Drop
```

---

### History Analytics

Charge and discharge history panels include:

- Current power
- Minimum power
- Maximum power
- Average power

Battery delta analysis:

- Δ1m
- Δ5m
- Δ10m

Session information:

- Duration
- Start %
- Current %
- Δ Session

---

## Example Critical Battery Drop

```text
Alerts & Events

Status          CRITICAL DROP

Drops Today     1
Spikes Today    4

Change          15.4 % → 5.2 %
Drop            -10.2 %
Window          60s
Mode            Discharging
Power           35.8 W
```

---

## Screenshots

### Main Dashboard

docs/images/dashboard.png

---

## Requirements

Python 3.11+

Install dependencies:

```bash
pip install rich psutil
```

---

## Running BatteryTop

```bash
python BatteryTop_v9_2_AlertsEventsWarnings.py
```

---

## Logging

BatteryTop automatically stores CSV logs.

### Battery Samples

```text
BatteryLogs/
```

### Power Spikes

```text
BatteryTop_PowerSpikes_YYYY-MM-DD.csv
```

### Battery Drops

```text
BatteryTop_BatteryDrops_YYYY-MM-DD.csv
```

---

## Battery Drop Rules

Default configuration:

```python
BATTERY_DROP_THRESHOLD_PERCENT = 3.0
BATTERY_DROP_WINDOW_SECONDS = 60

BATTERY_DROP_CRITICAL_PERCENT = 10.0
```

Examples:

### Battery Drop

```text
12.0 % → 8.0 %
```

Detected as:

```text
BATTERY DROP
```

### Critical Drop

```text
15.4 % → 5.2 %
```

Detected as:

```text
CRITICAL DROP
```

---

## Project Status

Current Version:

```text
v9.2
```

Implemented:

- Runtime Prediction
- Battery Health Analytics
- Charge Session Tracking
- Power Spike Detection
- Battery Drop Detection
- Alerts & Events Dashboard
- CSV Logging
- History Analytics

---

## Roadmap

Future ideas:

- Export reports
- Battery trend forecasting
- Battery degradation analysis
- Multi-battery support
- Theme customization

---

## License

MIT License
