# BatteryTop v10.3.2

A real-time battery analytics dashboard for Windows laptops that goes far beyond the standard Windows battery indicator.

BatteryTop provides deep insight into:

- Battery health
- Charging efficiency
- Discharge behavior
- Runtime prediction
- Battery wear
- Power spikes
- Battery percentage anomalies
- Charge and discharge session analytics
- Historical battery behavior

The goal of BatteryTop is not only to monitor battery status, but to help users understand how their laptop battery behaves over time and identify degradation, inefficient charging, abnormal discharge patterns and battery calibration issues. 【1-5e7202】

---

# Features

## Real-Time Battery Dashboard

Updates every second using:

```python
REFRESH_SECONDS = 1
```

Battery data is collected directly from Windows WMI providers.

BatteryTop continuously monitors:

- Battery percentage
- Voltage
- Current
- Charge power
- Discharge power
- Capacity
- Battery health
- Runtime estimates
- Session statistics
- Historical trends

---

# Dashboard Overview

## Battery Gauge

Vertical battery indicator.

### Purpose

Provides an intuitive visualization of current battery charge.

### Color Logic

| Battery % | Color |
|------------|--------|
| > 30% | Green |
| 15 - 30% | Yellow |
| 5 - 15% | Red |
| < 5% | Bright Red |

### Display

- 100% at top
- 0% at bottom
- Dynamic fill level

The gauge updates continuously and scales with battery charge. 【1-5e7202】

---

# Battery Panel

Provides live battery telemetry.

## Battery Level

Current battery charge percentage.

Formula:

```text
Battery % = Remaining Capacity / Full Capacity × 100
```

---

## Voltage

Current battery voltage.

Raw WMI value:

```text
Millivolts
```

Converted to:

```python
voltage_v = millivolts / 1000
```

---

## Charge Power

Current charging power.

Formula:

```text
Power (W) = Voltage × Current
```

Positive value during charging.

---

## Discharge Power

Current discharge rate.

Formula:

```text
Power (W) = Voltage × Current
```

Positive value while discharging.

---

## Average Power 1m / 5m

Moving averages.

### 1 Minute

Based on:

```python
power_history_1m
```

Maximum:

```python
60 samples
```

### 5 Minute

Based on:

```python
power_history_5m
```

Maximum:

```python
300 samples
```

Formula:

```python
average = sum(values) / len(values)
```

The averages smooth short-term fluctuations.

【1-5e7202】

---

# Battery Health

Provides information about long-term battery degradation.

---

## Design Capacity

Original battery capacity from manufacturer.

Example:

```text
94.946 Wh
```

---

## Current Capacity

Current maximum charge capacity.

Example:

```text
95.823 Wh
```

---

## Remaining

Current stored energy.

Example:

```text
87.773 Wh
```

---

## Wear Level

Formula:

```text
Wear %
=
(
Design Capacity - Full Charge Capacity ) / Design Capacity × 100
```

Example:

```text
Design: 95 Wh

Current: 85 Wh

Wear: 10.5 %
```

---

## Battery Health

Formula:

```text
Health % = Current Capacity / Design Capacity × 100
```

Example:

```text
90 Wh / 100 Wh = 90 %
```

---

# Runtime Prediction

One of BatteryTop's most advanced functions.

Provides ETA until:

- Empty
- Full Charge

depending on current battery mode.

---

## Current ETA

Uses current instantaneous power.

Formula:

```text
ETA = Remaining Energy / Current Power
```

Example:

```text
90 Wh / 30 W = 3 hours
```

Instantaneous but noisy.

---

## 1 Minute ETA

Uses:

```python
avg_power_1m
```

Formula:

```text
ETA_1m = Remaining Energy / Average Power 1m
```

More stable.

---

## 5 Minute ETA

Uses:

```python
avg_power_5m
```

Formula:

```text
ETA_5m = Remaining Energy / Average Power 5m
```

Most stable power-based estimate.

---

## Rate ETA

Uses battery percentage change.

Formula:

```text
Rate (%/h) = Battery % Difference / Elapsed Time
```

Example:

```text
5% in 30 min  = 10%/h
```

Then:

```text
Time Remaining = Current % / Rate
```

or

```text
Time To Full = (100 - Current %) / Rate
```

---

## Predicted ETA

BatteryTop combines multiple models.

Inputs:

```text
Current ETA
1m ETA
5m ETA
Rate ETA
```

Weighted model:

```python
1m  = 25%
5m  = 50%
Rate = 25%
```

The result is much more stable than a single estimate.

【1-5e7202】

---

## Variance

Measures disagreement between ETA models.

Small variance:

```text
Reliable estimate
```

Large variance:

```text
Unstable estimate
```

---

## Confidence

Based on:

- Sample count
- Variance

Possible values:

```text
High
Medium
Low
```

---

## Samples

Number of historical measurements available.

More samples:

```text
Higher prediction confidence
```

---

# Session Analytics

Tracks the currently active charge or discharge session.

---

## Session

Displays:

```text
Start % → Current % (Delta)
```

Example:

```text
91.685% → 94.872% (+3.187%)
```

---

## Duration

Elapsed session time.

Formula:

```text
Now
-
Session Start Time
```

---

## Average Power

Average charging or discharge power during session.

Formula:

```text
Average = Sum(Power Samples) / Sample Count
```

---

## Peak Power

Highest measured power.

Formula:

```python
max(session_powers)
```

---

## Min Power

Lowest measured power.

Formula:

```python
min(session_powers)
```

---

## Samples

Number of recorded samples in current session.

---

# Charge Session Log

Tracks historical sessions.

---

## Charge Sessions

Number of completed charge sessions today.

---

## Discharge Sessions

Number of completed discharge sessions today.

---

## Charge Gain

Total battery gained today.

Example:

```text
+25%
```

---

## Discharge Loss

Total battery consumed today.

Example:

```text
-40%
```

---

## Current Delta

Current session battery change.

---

# Power Monitor

Analyzes short-term power behavior.

---

## Current Power

Present power draw.

---

## Average 5 Minutes

Rolling average power.

Used for:

- ETA
- Spike Detection

---

## Factor

Formula:

```text
Current Power / Average 5m
```

Example:

```text
50W / 25W = 2.0x
```

---

## Power Spike Detection

Formula:

```text
Current Power / Average 5m > 2.5x
```

Default threshold:

```python
POWER_SPIKE_FACTOR_THRESHOLD = 2.5
```

Power spikes are stored in:

```python
power_spike_history
```

【1-5e7202】

---

# Alerts & Events

Provides anomaly detection.

---

## Battery Drop Detection

Detects sudden battery percentage drops.

Default rules:

```python
3%  within 60 seconds
```

---

## Critical Battery Drop

Detects severe drops.

Default rule:

```python
10% within 60 seconds
```

Useful for:

- Battery calibration issues
- Firmware bugs
- Faulty battery packs

---

## Spikes Today

Number of detected power spikes.

---

## Drops Today

Number of battery drop events.

---

# Battery Mode + Percent Timeline

Tracks battery behavior over time.

---

## Mode Line

Shows:

```text
Charge
Discharge
Idle
```

history.

---

## Battery Line

Displays percentage evolution.

Higher blocks:

```text
Higher battery charge
```

Lower blocks:

```text
Lower battery charge
```

---

## Session

Displays:

```text
Start % → Current % (Delta)
```

---

## Duration

Current timeline duration.

---

# Charge / Discharge Power History

Separate histories are maintained.

---

## Cur / Min / Max / Avg

Displays:

```text
Current Power
Minimum
Maximum
Average
```

for current mode.

---

## Delta 1 / 5 / 10 Minutes

Formula:

```text
Current %  - Battery % X Minutes Ago
```

Examples:

```text
Δ1m
Δ5m
Δ10m
```

Useful for:

- Detecting battery drain
- Comparing charge efficiency
- Identifying power anomalies

---

## Session

Displays battery start and current values.

---

## Duration

How long the current session has been active.

---

# CSV Logging

BatteryTop automatically logs data.

Location:

```text
BatteryLogs/
```

---

## Main Log

```text
BatteryTop_YYYY-MM-DD.csv
```

Contains:

- Battery %
- Voltage
- Current
- Power
- Capacity
- Runtime metrics

---

## Power Spikes

```text
BatteryTop_PowerSpikes_YYYY-MM-DD.csv
```

Contains:

- Timestamp
- Power
- Factor
- Battery %

---

## Battery Drops

```text
BatteryTop_BatteryDrops_YYYY-MM-DD.csv
```

Contains:

- Previous %
- Current %
- Drop %
- Window
- Power

---

# Why BatteryTop Is Different

Most battery tools show:

```text
Battery %
Time Remaining
```

BatteryTop shows:

✅ Battery wear  
✅ Runtime confidence  
✅ Power spikes  
✅ Battery drops  
✅ Charge efficiency  
✅ Session analytics  
✅ Historical behavior  
✅ Capacity degradation  
✅ Predictive runtime models  

BatteryTop helps users understand **why** their battery behaves the way it does, rather than only reporting the current charge level. 【1-5e7202】
