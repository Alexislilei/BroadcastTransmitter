# DTV YAML Generator — v0.9

A desktop tool (Tkinter GUI + CLI) that turns the merged Australian
broadcast-transmitter dataset into a per-area frequency YAML file.

## Contents of this release

| File | Description |
| --- | --- |
| `DTV_YAML_Generator.exe` | **Windows** standalone executable — no Python required. |
| `DTV_YAML_Generator_macos.zip` | **macOS** standalone executable — no Python required. Unzip, then run `DTV_YAML_Generator`. |
| `dtv_yaml_generator.py` | Python source (run with `python dtv_yaml_generator.py`). |
| `BroadcastTransmitter_dtv_simplified_merged.json` | Dataset (required by the Python script). |
| `callsign_name_lookup.json` | Callsign → name lookup table (required by the Python script). |

> The two JSON files are **already bundled inside** the `.exe` and the macOS
> executable, so the standalone binaries run on their own. They are shipped
> separately only for the Python-script workflow.

## Running the standalone executables

### Windows
Double-click `DTV_YAML_Generator.exe`. The GUI opens.

### macOS
```bash
unzip DTV_YAML_Generator_macos.zip
chmod +x DTV_YAML_Generator
./DTV_YAML_Generator
```
Because the binary is unsigned, macOS Gatekeeper may block it on first launch.
If so, right-click the file → **Open**, or run:
```bash
xattr -d com.apple.quarantine DTV_YAML_Generator
```

## Running from source (Python 3.9+)

```bash
python dtv_yaml_generator.py
```

Keep `dtv_yaml_generator.py`, `BroadcastTransmitter_dtv_simplified_merged.json`
and `callsign_name_lookup.json` in the same directory.

## Usage

1. Pick a **State** from the first drop-down.
2. Pick an **Area Served** from the second drop-down.
3. Click **Generate YAML** and choose where to save the file.

The output file is named after the area, e.g. `Gold Coast.yaml`.

### Command line

```bash
# List all states and areas
DTV_YAML_Generator --list

# Generate a YAML file for a specific area
DTV_YAML_Generator --state QLD --area "Gold Coast" --out "Gold Coast.yaml"

# Print the nearest transmitter site to the current location
DTV_YAML_Generator --nearest
```

## Output format

```yaml
frequencies:
  - channel: 40
    frequency_hz: 613500000
    frequency_mhz: 613.5
    bandwidth_mhz: 7
    standard: DVB-T
    Callsign: SBS
    FullName: SBS_Television
    network_hint: SBS
    source_name: CH40_6135_SBS
```

`FullName` and `network_hint` are looked up from `callsign_name_lookup.json`
using the callsign prefix (e.g. `TNQ40` → `TNQ`). `source_name` is built as
`CH<channel>_<frequency>_<network_hint>`.
