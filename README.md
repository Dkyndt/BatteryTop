# BatteryTop

BatteryTop is an advanced real-time battery monitoring dashboard for Windows laptops.

Built with Python and Rich, BatteryTop provides battery telemetry, health analytics, runtime prediction, session tracking, power spike detection and battery drop alerts in a compact terminal interface.

---

## Features

### Battery Monitoring

- Battery percentage
- Voltage
- Current
- Charge power
- Discharge power
- Remaining energy
- Full capacity
- Cycle count
- Estimated runtime

---

### Battery Health

- Design capacity
- Current capacity
- Remaining energy
- Wear level

---

### Runtime Prediction

Multiple prediction models:

- Current battery estimate
- 1 minute power model
- 5 minute power model

Provides a more stable runtime estimate than the standard Windows battery prediction.

---

### Charge and Discharge Analytics

Tracks:

- Current session
- Session duration
- Start percentage
- Current percentage
- Percentage change

---

### Charge Session Log

Tracks:

- Charge sessions
- Discharge sessions
- Charge gain
- Discharge loss

---

### Power Monitor

Displays:

- Current power draw
- 5 minute average
- Spike factor
- Recent power spikes

---

### Power Spike Detection

Automatically detects unusual power consumption peaks.

Examples:

- Starting Visual Studio
- Large compilations
- Teams meetings
- OneDrive sync operations

Power spikes are logged and displayed in the dashboard.

---

### Battery Drop Detection

Detects sudden battery percentage drops.

Example:

15.4% → 5.2%

BatteryTop identifies this as:

CRITICAL BATTERY DROP

with:

- detected percentage change
- timestamp
- battery mode
- power consumption

---

### Alerts & Events

Central event dashboard showing:

- Power spikes
- Battery drops
- Critical battery drops
- Daily statistics

---

### Battery Timeline

Shows:

- Charging periods
- Discharging periods
- Battery percentage history

---

### History Analytics

Both charge and discharge panels display:

- Current power
- Minimum power
- Maximum power
- Average power

Battery delta tracking:

- Δ1m
- Δ5m
- Δ10m

Session information:

- Duration
- Start %
- Current %
- Δ Session

---

## Screenshots

### BatteryTop Dashboard

Add screenshot here:

docs/images/dashboard.png
