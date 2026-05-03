# Hardware Evaluation

Automated hardware and software assessment for Windows laptops and desktops. Runs a full diagnostic across 12 modules, scores each component, and generates a self-contained HTML report with a buy/no-buy recommendation.

## Quick Start

```powershell
irm https://raw.githubusercontent.com/maxdorx/hardware-eval/main/Hardware-Evaluation.ps1 | iex
```

> Run in an elevated (Administrator) PowerShell window for full results. The script will offer to re-launch elevated automatically if needed.

---

## What It Tests

| Module | What's checked |
|---|---|
| **CPU** | Core count, generation, tier (i3/i5/i7/Ryzen), idle/load temperature, WinSAT score |
| **RAM** | Capacity, speed (MHz), type (DDR4/5), single-channel detection, usage pressure |
| **Storage** | Drive type (NVMe/SSD/HDD), SMART health via smartctl, read/write speed via DiskSpd, capacity |
| **Battery** | Health %, full-charge capacity (Wh), cycle count, charge rate |
| **GPU** | Discrete vs integrated, VRAM, driver version, display resolution and refresh rate |
| **Network** | Wi-Fi generation (Wi-Fi 5/6/6E), download speed via Speedtest CLI, ping, Ethernet presence |
| **Security** | TPM version, Secure Boot, BitLocker, Windows Defender / AV status, firewall |
| **Ports** | USB 3.x count, USB-C, Thunderbolt, HDMI/DP, SD reader, Bluetooth, fingerprint reader |
| **Software** | Windows activation, pending updates, driver errors, BSOD history, startup program count, dirty volumes |
| **Thermal** | CPU and GPU temps at idle and under load (via LibreHardwareMonitor when available) |
| **Audio** | Controllers, playback and recording devices, microphone level test, driver errors |
| **Input** | Keyboard and touchpad detection, driver health |

---

## Output

The script prints a live summary to the terminal and saves a **self-contained HTML report** to `Desktop\Reports\` (configurable).

The report includes:
- Overall score (0–100) with a verdict
- Per-module scores with expandable detail cards
- Department suitability scores (General Use, Software Dev, Design, Video Editing, IT Dept, etc.)
- Full diagnostic log

**Verdicts**

| Score | Verdict |
|---|---|
| 90–100 | Highly Recommended |
| 78–89 | Recommended |
| 65–77 | Conditionally Recommended |
| 50–64 | Use Case Dependent |
| < 50 | Not Recommended |

---

## Requirements

- Windows 10 (Build 19041+) or Windows 11
- PowerShell 5.1 or later
- Internet access on first run (downloads tools; cached afterwards)
- Administrator privileges recommended

---

## Parameters

```powershell
# Fully automatic — no prompts, runs all modules
.\Hardware-Evaluation.ps1 -FullAuto

# Specific modules only
.\Hardware-Evaluation.ps1 -TestModules "CPU,RAM,Storage"

# Save report to a custom location
.\Hardware-Evaluation.ps1 -OutputPath "C:\Reports"

# No internet — skip tool downloads, use Windows built-ins only
.\Hardware-Evaluation.ps1 -NoDownload

# Collect results without generating the HTML file
.\Hardware-Evaluation.ps1 -NoReport

# Load a custom JSON config (override thresholds, weights, company name)
.\Hardware-Evaluation.ps1 -ConfigPath ".\my-config.json"
```

---

## Custom Configuration

Override any threshold or scoring weight with a JSON file:

```json
{
  "Thresholds": {
    "RAM_MinGB": 16,
    "RAM_RecommendedGB": 32,
    "Storage_MinGB": 512,
    "Battery_MinHealth": 75
  },
  "Weights": {
    "CPU": 25,
    "RAM": 20,
    "Storage": 20,
    "Battery": 15,
    "GPU": 5,
    "Network": 5,
    "Security": 5,
    "Software": 5
  },
  "Report": {
    "CompanyName": "Acme IT"
  }
}
```

Pass it with `-ConfigPath ".\config.json"`. Any key not present falls back to the built-in default.

---

## Tools Downloaded

On first run the script downloads these tools into `%LOCALAPPDATA%\HWEval\tools\` and reuses them on subsequent runs:

| Tool | Purpose |
|---|---|
| [DiskSpd](https://github.com/microsoft/diskspd) | Sequential read/write benchmark |
| [Speedtest CLI](https://www.speedtest.net/apps/cli) | Download/upload speed and ping |
| [smartmontools](https://www.smartmontools.org/) | Full SMART attribute data per drive |

Use `-NoDownload` to skip all downloads and fall back to Windows built-in equivalents (WMI SMART, WinSAT scores, CDN speed estimate).

---

## Elevation

The script works without admin rights but the following are limited or unavailable without elevation:

- SMART data (requires kernel access)
- WinSAT scores
- Thermal sensors (MSAcpi_ThermalZoneTemperature)
- BitLocker status
- Battery cycle count (some hardware)

The script will prompt to re-launch elevated, or do so silently if `-FullAuto` is set.
