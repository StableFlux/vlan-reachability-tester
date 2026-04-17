# VLAN Reachability Tester

<p align="center">
  <img src="logo.png" alt="VLAN Reachability Tester" width="180"/>
</p>

A portable Windows 11 desktop application for testing network reachability across VLANs in real time. No installation required — just download and run.

![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-GPLv3-blue)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-stableflux-orange?logo=buy-me-a-coffee)](https://buymeacoffee.com/stableflux)

---

## Screenshots

![Monitor Tab](monitor.png)
*Monitor tab — live reachability matrix and current sweep status*

![Config Tab](config.png)
*Config tab — VLAN definitions, NIC selection, and ping settings*

---

## Features

- **Live reachability matrix** — colour-coded grid showing reachability between all configured VLANs
- **Per-sweep status table** — shows current sweep results with RTT, target IP, and device label
- **NIC selection** — bind all testing to a specific network adapter
- **IP renewal** — release and renew DHCP on the selected adapter without leaving the app
- **Persistent results** — matrix retains history across sweeps with a manual clear button
- **Export** — save results to `vlan_results.json` and `vlan_matrix.txt`
- **Dark theme GUI** — built with tkinter, no external dependencies at runtime
- **Fully portable** — single `.exe`, no Python installation needed

---

## Download

Download the latest `VLAN Tester.exe` from the [Releases](https://github.com/StableFlux/vlan-reachability-tester/releases) page.

---

## Getting Started

### 1. Run the application
Double-click `VLAN Tester.exe` — no installation needed. On first run a `vlan_config.json` file is created in the same folder to store your configuration.

### 2. Configure your VLANs
Click the **Config** tab.

Under **VLAN Definitions**, click **＋ Add** to add each VLAN:

| Field | Description | Example |
|-------|-------------|---------|
| VLAN Name | Short identifier | `TRUSTED` |
| Subnet | Network prefix with trailing dot | `10.10.20.` |
| Target IP | Device to ping on that VLAN | `10.10.20.1` |
| Device Label | Friendly name for the target | `Core Switch` |

Repeat for each VLAN you want to test. Use the **↑ / ↓** buttons to reorder.

### 3. Select your network adapter
Under **Network Interface**, click the adapter you want to use for testing. The selected adapter is highlighted in green. Click **↺ Refresh** if your adapter is not listed.

### 4. Configure ping settings

| Setting | Description | Default |
|---------|-------------|---------|
| Sweep interval | Seconds between full sweeps | `5` |
| Ping timeout | Seconds to wait per ping | `2` |
| Pings per host | Number of pings sent per target | `1` |

### 5. Apply and start
Click **✔ Apply & Restart Sweep** in the tab bar. The app switches to the **Monitor** tab and begins sweeping.

---

## Monitor Tab

### Status bar
Shows your current VLAN, IP address, selected NIC, and sweep count — updated every second even while paused.

### Current VLAN Sweep table
Live results for the current sweep:

| Status | Meaning |
|--------|---------|
| `REACH` | Target is reachable |
| `BLOCK` | Target did not respond |
| `THIS SUBNET` | This is the VLAN your adapter is currently on |
| `???` | Not yet tested this sweep |

### Full Reachability Matrix
Persistent grid showing historical results for every source→destination VLAN pair. Results are retained across sweeps until you click **🗑 CLEAR MATRIX**.

| Colour | Meaning |
|--------|---------|
| Green | Reachable |
| Red | Blocked |
| Grey | Not yet tested |

### Toolbar buttons

| Button | Action |
|--------|--------|
| **⏸ PAUSE / ▶ RESUME** | Pause or resume the sweep |
| **🔄 RENEW IP** | Pauses sweep, runs ipconfig release/renew on selected adapter, waits for new IP |
| **🗑 CLEAR MATRIX** | Clears the reachability matrix and resets sweep count |
| **💾 EXPORT** | Saves results to `vlan_results.json` and `vlan_matrix.txt` |

---

## Moving Between VLANs

When you physically move your network connection to a different VLAN (cable change, different Wi-Fi SSID, or switch port reconfiguration):

1. Click **🔄 RENEW IP** — the app pauses the sweep, releases and renews the DHCP lease on your selected adapter, and waits for the new IP
2. Once the new IP appears the status bar updates automatically
3. Click **▶ RESUME** to continue sweeping from the new VLAN

---

## Subnet Format

Subnets must be entered with a **trailing dot**:

| Correct | Incorrect |
|---------|-----------|
| `10.10.20.` | `10.10.20.0/24` |
| `192.168.10.` | `192.168.10.0` |

---

## Configuration File

Settings are saved automatically to `vlan_config.json` in the same folder as the exe. You can copy this file alongside the exe to carry your configuration to another machine.

---

## Building from Source

Requires Python 3.10+ and PyInstaller.

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "VLAN Tester" vlan_tester_gui.py
```

The compiled exe will be in the `dist/` folder.

---

## Raspberry Pi Version

A terminal-based version for Raspberry Pi is included at `Raspberry Pi/vlan_tester_pi.py`. It uses Linux ping flags (`-c`, `-W`) and requires root. Intended to be run directly in the Pi shell.

---

## Requirements

**To run the exe:**
- Windows 10 or 11 (64-bit)
- No Python installation needed

**To run from source:**
- Python 3.10+
- No third-party packages required (tkinter is included with Python)

---

## License

GNU General Public License v3.0 — free to use, modify, and distribute, but any derivative works must also be open source under the same licence. Commercial resale is not permitted.
