<div align="center">

<img src="Araneo_inventorytr.png" alt="Araneo Inventory Icon" width="120"/>

# Araneo Inventory

**Export a full network inventory from Luminex Araneo `.ara` project files to Excel**

[![Version](https://img.shields.io/badge/version-1.0.5-blue)](#)
[![Python](https://img.shields.io/badge/python-3.8%2B-blue?logo=python&logoColor=white)](#)
[![Platform](https://img.shields.io/badge/platform-Windows-blue?logo=windows)](#)
[![License](https://img.shields.io/badge/license-MIT-green)](#)

</div>

---

## Overview

**Araneo Inventory** reads `.ara` project files saved by the [Luminex Araneo](https://www.luminex.be/products/software/araneo/) network management software and exports a structured, colour-coded Excel workbook containing a complete inventory of all switches, ports, VLAN groups, and connected devices.

---

## Excel Output Structure

| Tab | Contents |
|---|---|
| **Summary** | Project metadata · all switches (clickable) · all groups/VLANs (clickable, colour-coded) |
| **`<Switch Name>`** | Switch info block · full port table with group colours, connected device, SFP data |
| **`G_<Group Name>`** | Group info · VLAN ID · device table: IP · MAC · Vendor · Label · Switch (link) · Port # · Port Name |

### Switch tab — port table columns

| Port # | Name | Enabled | Group ID | Group Name | Connected MAC | Vendor | IP Address | *(SFP columns when present)* |
|---|---|---|---|---|---|---|---|---|

### Group tab — device table columns

| IP Address | MAC Address | Vendor | Device / Label | Switch | Port # | Port Name |
|---|---|---|---|---|---|---|

---

## Requirements

### Running from source

- Python 3.8 or newer
- `openpyxl` — Excel file generation
- `Pillow` — icon processing (only needed at build time)

```bash
pip install openpyxl pillow
python araneo_inventory.py
```

### Building the Windows executable

- Python 3.8+ with internet access (first run only)
- Double-click **`build_exe.bat`** — it installs dependencies and runs PyInstaller automatically

```
Araneo_Inventory.exe   ← standalone, no Python required on target PC
```

---

## Usage

### GUI (double-click the `.exe` or run the script)

1. **Input .ara file** — browse to your Araneo project file
2. **Output .xlsx file** — pre-filled automatically in the same folder as the source file
3. **Project name** — stored in the Summary header (optional)
4. **Creator** — stored in the Summary header (optional)
5. **Date created** — filled automatically from the system clock
6. Click **⚙ Extract Inventory**

### Command line

```bash
# Basic usage — output saved next to the .ara file
python araneo_inventory.py project.ara

# Specify output path, project name, and creator
python araneo_inventory.py project.ara output.xlsx "Show Name" "Your Name"

# Same with the compiled exe
Araneo_Inventory.exe project.ara output.xlsx "Show Name" "Your Name"
```

---

## Build

```
project/
├── araneo_inventory.py   # Main script — GUI + extraction logic, icon embedded as base64
├── araneo_inventory.ico  # Multi-resolution icon for PyInstaller
├── build_exe.bat         # One-click Windows EXE builder
└── README.md
```

To build the standalone executable on Windows:

```bat
build_exe.bat
```

This installs `openpyxl`, `pyinstaller`, and `pillow` via pip, then produces `Araneo_Inventory.exe` in the current folder.

---




## License

MIT — see `LICENSE` for details.

---

<div align="center">
Made for live-event network engineers using <a href="https://www.luminex.be">Luminex</a> GigaCore infrastructure.
</div>
