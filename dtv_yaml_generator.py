#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DTV YAML Generator
==================

A small desktop tool (Tkinter GUI) that turns the merged broadcast-transmitter
dataset (``BroadcastTransmitter_dtv_simplified_merged.json``) into a per-area
frequency YAML file, in the same format as ``Goldcoast_示例.yaml`` /
``Kiama_示例.yaml``.

Usage
-----
    python dtv_yaml_generator.py

Then:
    1. Pick a State from the first drop-down.
    2. Pick an Area Served from the second drop-down.
    3. Click "Generate YAML" and choose where to save the file.

The output file is named after the area, e.g. ``Gold Coast.yaml``.

The GUI also offers an "Open mySwitch" button that launches the default web
browser at the official Australian DTV coverage lookup page
(``https://myswitch.digitalready.gov.au/``). An optional address entered in
the adjacent box is pre-filled into the mySwitch search field.

On startup the GUI tries to determine the current geographic location and
preselects the nearest transmitter site. Because straight-line distance does
not always match the official broadcast service area (which is defined by
licensed coverage, not proximity), the five closest sites are listed with
their distances so the user can confirm or pick another one. Location is
resolved in this order:

1. The Windows Location Service (``Windows.Devices.Geolocation``), which
   fuses GPS, Wi-Fi hotspot databases and IP - the same class of data Google
   Maps uses, and far more accurate than IP alone.
2. Free IP-based geolocation services (no API key required). Note that IP
   geolocation reflects the ISP's registered address for the IP block, so it
   can be off by tens of kilometres.
3. If neither works, the first record in the dataset is selected.

The program can also be used from the command line (no GUI):

    python dtv_yaml_generator.py --list
    python dtv_yaml_generator.py --state QLD --area "Gold Coast" --out "Gold Coast.yaml"
    python dtv_yaml_generator.py --nearest
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import subprocess
import sys
import urllib.parse
import urllib.request
import webbrowser
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

#: Default dataset file (relative to this script / current working directory).
DEFAULT_DATA_FILE = "BroadcastTransmitter_dtv_simplified_merged.json"

#: Country code written into the YAML header.
COUNTRY = "AU"

#: Official Australian DTV signal coverage lookup page (mySwitch).
MYSWITCH_URL = "https://myswitch.digitalready.gov.au/"

#: Recording duration written into the YAML header (seconds).
RECORD_DURATION_SECONDS = 90

#: DVB-T channel bandwidth in MHz (Australia uses 7 MHz channels).
BANDWIDTH_MHZ = 7

#: Broadcast standard written into each frequency entry.
STANDARD = "DVB-T"

#: Default callsign -> name lookup table (relative to this script / cwd).
#: The table is a JSON array of objects with the keys ``callsign``,
#: ``fullName`` and ``name``, e.g.::
#:
#:     [{"callsign": "TNQ", "fullName": "10_Regional_Queensland", "name": "Ten"}, ...]
#:
#: ``name`` becomes the ``network_hint`` and ``fullName`` becomes ``FullName``.
DEFAULT_LOOKUP_FILE = "callsign_name_lookup.json"

#: Mean Earth radius in kilometres (used for great-circle distances).
EARTH_RADIUS_KM = 6371.0088

#: IP-based geolocation endpoints, tried in order. Each must return JSON with
#: ``lat`` and ``lon`` (or ``latitude``/``longitude``) numeric fields.
GEOLOCATION_ENDPOINTS: Tuple[str, ...] = (
    "http://ip-api.com/json/?fields=status,lat,lon,city,regionName,country",
    "https://ipapi.co/json/",
    "https://ipinfo.io/json",
)

#: Timeout (seconds) for each geolocation request.
GEOLOCATION_TIMEOUT = 4.0

#: Timeout (seconds) for the Windows system location service.
SYSTEM_LOCATION_TIMEOUT = 12.0

#: PowerShell script that queries the Windows Location Service
#: (``Windows.Devices.Geolocation``). This uses the OS location stack, which
#: fuses GPS, Wi-Fi hotspot databases and IP, and is therefore far more
#: accurate than IP-only geolocation. Prints ``LAT=<lat> LON=<lon> ACC=<m>``
#: on success and exits non-zero on failure.
_SYSTEM_LOCATION_PS = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$type = [Windows.Devices.Geolocation.Geolocator,Windows.Devices.Geolocation,ContentType=WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
function Await($op, $t) {
    $m = $asTask.MakeGenericMethod($t)
    $task = $m.Invoke($null, @($op))
    if (-not $task.Wait(10000)) { throw 'timeout' }
    return $task.Result
}
$loc = [Activator]::CreateInstance($type)
$loc.DesiredAccuracyInMeters = 100
$pos = Await ($loc.GetGeopositionAsync()) ([Windows.Devices.Geolocation.Geoposition,Windows.Devices.Geolocation,ContentType=WindowsRuntime])
$c = $pos.Coordinate.Point.Position
Write-Output ('LAT=' + $c.Latitude + ' LON=' + $c.Longitude + ' ACC=' + $pos.Coordinate.Accuracy)
"""


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_dataset(path: str) -> Dict[str, Dict[str, Any]]:
    """Load the merged JSON dataset.

    The expected structure is::

        {
          "QLD": {
            "Gold Coast": {
              "Site id": 403573,
              "Site Name": "...",
              "Latitude": "...",
              "Longitude": "...",
              "Transmissions": [
                {"Callsign": "SBS40", "Frequency(MHz)": 613.5, "Channel": "40"},
                ...
              ]
            },
            ...
          },
          ...
        }

    Returns the parsed dictionary. Raises ``FileNotFoundError`` if the file
    does not exist and ``ValueError`` if the structure is not as expected.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Dataset file not found: {path}")

    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    if not isinstance(data, dict):
        raise ValueError(
            "Unexpected dataset structure: top level must be an object "
            "keyed by state."
        )
    return data


def list_states(data: Dict[str, Dict[str, Any]]) -> List[str]:
    """Return the sorted list of states present in the dataset."""
    return sorted(data.keys())


def list_areas(data: Dict[str, Dict[str, Any]], state: str) -> List[str]:
    """Return the sorted list of areas served for a given state."""
    areas = data.get(state, {})
    if not isinstance(areas, dict):
        return []
    return sorted(areas.keys())


def load_lookup_table(path: str) -> Dict[str, Dict[str, str]]:
    """Load the callsign -> name lookup table.

    The table is a JSON array of objects with the keys ``callsign``,
    ``fullName`` and ``name``, e.g.::

        [{"callsign": "TNQ", "fullName": "10_Regional_Queensland", "name": "Ten"}, ...]

    Returns a dictionary keyed by the upper-cased callsign, whose values are
    ``{"fullName": ..., "name": ...}``. A missing file yields an empty table
    (the generator then falls back to the callsign prefix for both fields).
    """
    if not os.path.isfile(path):
        return {}

    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)

    table: Dict[str, Dict[str, str]] = {}
    if not isinstance(raw, list):
        return table

    for row in raw:
        if not isinstance(row, dict):
            continue
        callsign = str(row.get("callsign", "") or "").strip().upper()
        if not callsign:
            continue
        table[callsign] = {
            "fullName": str(row.get("fullName", "") or "").strip(),
            "name": str(row.get("name", "") or "").strip(),
        }
    return table


# ---------------------------------------------------------------------------
# Geolocation helpers
# ---------------------------------------------------------------------------

def parse_dms(value: Any) -> Optional[float]:
    """Parse a DMS coordinate string into decimal degrees.

    Handles the dataset format ``"35 16 32S"`` / ``"149 5 52E"`` as well as
    plain decimal strings such as ``"-35.2755"``. Returns ``None`` when the
    value cannot be interpreted.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    # Plain decimal number (possibly signed).
    try:
        return float(text)
    except ValueError:
        pass

    # DMS form: degrees [minutes [seconds]] [N/S/E/W]
    match = re.match(
        r"^\s*([+-]?\d+(?:\.\d+)?)\s*[°\s]\s*"
        r"(?:(\d+(?:\.\d+)?)\s*['\s]\s*)?"
        r"(?:(\d+(?:\.\d+)?)\s*[\"\s]?\s*)?"
        r"([NSEWnsew])?\s*$",
        text,
    )
    if not match:
        return None

    degrees = float(match.group(1))
    minutes = float(match.group(2)) if match.group(2) else 0.0
    seconds = float(match.group(3)) if match.group(3) else 0.0
    hemisphere = (match.group(4) or "").upper()

    decimal = abs(degrees) + minutes / 60.0 + seconds / 3600.0
    if degrees < 0 or hemisphere in ("S", "W"):
        decimal = -decimal
    return decimal


def haversine_km(
    lat1: float, lon1: float, lat2: float, lon2: float
) -> float:
    """Great-circle distance between two points in kilometres."""
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = (
        math.sin(d_phi / 2.0) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2.0) ** 2
    )
    return 2.0 * EARTH_RADIUS_KM * math.asin(math.sqrt(a))


def get_system_location(
    timeout: float = SYSTEM_LOCATION_TIMEOUT,
) -> Optional[Tuple[float, float]]:
    """Return ``(latitude, longitude)`` from the Windows Location Service.

    This queries ``Windows.Devices.Geolocation`` via PowerShell, which fuses
    GPS, Wi-Fi hotspot databases and IP to produce a much more accurate fix
    than IP-only geolocation. Returns ``None`` on non-Windows platforms, when
    the location service is disabled, or when the query fails/times out.
    """
    if os.name != "nt":
        return None

    try:
        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                _SYSTEM_LOCATION_PS,
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except Exception:  # noqa: BLE001 - missing powershell, timeout, etc.
        return None

    if completed.returncode != 0:
        return None

    match = re.search(
        r"LAT=(-?\d+(?:\.\d+)?)\s+LON=(-?\d+(?:\.\d+)?)", completed.stdout
    )
    if not match:
        return None

    lat = float(match.group(1))
    lon = float(match.group(2))
    if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
        return lat, lon
    return None


def get_ip_location(
    timeout: float = GEOLOCATION_TIMEOUT,
) -> Optional[Tuple[float, float]]:
    """Return ``(latitude, longitude)`` from free IP-based geolocation.

    No API key is required. Note that IP geolocation reflects the ISP's
    registered address for the IP block (often a data centre or regional
    aggregation point), so it can be off by tens of kilometres. Returns
    ``None`` if every endpoint fails or the network is unavailable.
    """
    for url in GEOLOCATION_ENDPOINTS:
        try:
            request = urllib.request.Request(
                url, headers={"User-Agent": "dtv-yaml-generator/1.0"}
            )
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = json.loads(response.read().decode("utf-8", "replace"))
        except Exception:  # noqa: BLE001 - try the next endpoint
            continue

        if not isinstance(payload, dict):
            continue
        # ip-api.com reports failures via a "status" field.
        if str(payload.get("status", "")).lower() == "fail":
            continue

        lat = payload.get("lat", payload.get("latitude"))
        lon = payload.get("lon", payload.get("longitude", payload.get("lng")))
        try:
            lat_f = float(lat)
            lon_f = float(lon)
        except (TypeError, ValueError):
            continue

        if -90.0 <= lat_f <= 90.0 and -180.0 <= lon_f <= 180.0:
            return lat_f, lon_f

    return None


def get_current_location(
    timeout: float = GEOLOCATION_TIMEOUT,
) -> Optional[Tuple[float, float]]:
    """Return the current ``(latitude, longitude)`` in decimal degrees.

    Tries the most accurate source first and falls back progressively:

    1. Windows Location Service (GPS / Wi-Fi / IP fusion) - most accurate.
    2. Free IP-based geolocation services - coarse, city-level at best.

    Returns ``None`` when no source could provide a location.
    """
    location = get_system_location()
    if location is not None:
        return location
    return get_ip_location(timeout=timeout)


def find_nearest_sites(
    data: Dict[str, Dict[str, Any]],
    latitude: float,
    longitude: float,
    limit: int = 5,
) -> List[Tuple[str, str, float]]:
    """Return the ``limit`` transmitter sites closest to the coordinates.

    Each item is ``(state, area, distance_km)``, sorted nearest first. Sites
    without usable coordinates are skipped. Returns an empty list when no
    site in the dataset has usable coordinates.

    Note: straight-line distance does not necessarily match the official
    broadcast service area (which is defined by licensed coverage, not
    proximity), so callers should let the user confirm the selection.
    """
    candidates: List[Tuple[str, str, float]] = []
    for state, areas in data.items():
        if not isinstance(areas, dict):
            continue
        for area, site in areas.items():
            if not isinstance(site, dict):
                continue
            site_lat = parse_dms(site.get("Latitude"))
            site_lon = parse_dms(site.get("Longitude"))
            if site_lat is None or site_lon is None:
                continue
            distance = haversine_km(latitude, longitude, site_lat, site_lon)
            candidates.append((state, area, distance))

    candidates.sort(key=lambda item: item[2])
    return candidates[: max(0, limit)]


def find_nearest_site(
    data: Dict[str, Dict[str, Any]],
    latitude: float,
    longitude: float,
) -> Optional[Tuple[str, str, float]]:
    """Find the transmitter site closest to the given coordinates.

    Returns ``(state, area, distance_km)`` or ``None`` when no site in the
    dataset has usable coordinates.
    """
    nearest = find_nearest_sites(data, latitude, longitude, limit=1)
    return nearest[0] if nearest else None


def first_record(data: Dict[str, Dict[str, Any]]) -> Optional[Tuple[str, str]]:
    """Return the first ``(state, area)`` record in the dataset, if any."""
    for state in list_states(data):
        areas = list_areas(data, state)
        if areas:
            return state, areas[0]
    return None


# ---------------------------------------------------------------------------
# mySwitch web lookup
# ---------------------------------------------------------------------------

def build_myswitch_url(address: Optional[str] = None) -> str:
    """Build a mySwitch URL, optionally pre-filling the address search box.

    mySwitch (``https://myswitch.digitalready.gov.au/``) is the official
    Australian DTV signal coverage lookup tool. It accepts an ``address``
    query parameter that pre-populates its search field, e.g.::

        https://myswitch.digitalready.gov.au/?address=1%20George%20St%20Sydney

    Spaces are percent-encoded as ``%20`` (not ``+``). ``urlencode`` would
    turn spaces into ``+``, but mySwitch reads the raw parameter value without
    decoding ``+`` back to a space, so the search box would literally show
    ``8+lightwood+ct``. Using ``quote`` keeps the address readable.

    When ``address`` is empty or ``None`` the plain landing page is returned.
    """
    if not address or not address.strip():
        return MYSWITCH_URL
    query = urllib.parse.quote(address.strip(), safe="")
    return f"{MYSWITCH_URL}?address={query}"


def open_myswitch(address: Optional[str] = None) -> bool:
    """Open the mySwitch coverage page in the default web browser.

    Returns ``True`` when the browser was launched successfully, ``False``
    otherwise. The optional ``address`` is passed through to
    :func:`build_myswitch_url` so the search box can be pre-filled.
    """
    url = build_myswitch_url(address)
    try:
        return webbrowser.open(url, new=2)
    except Exception:  # noqa: BLE001 - browser launch is best-effort
        return False


# ---------------------------------------------------------------------------
# Transformation helpers
# ---------------------------------------------------------------------------

def callsign_prefix(callsign: str) -> str:
    """Strip trailing digits from a callsign.

    ``"SBS40"`` -> ``"SBS"``, ``"BTQ42"`` -> ``"BTQ"``, ``"ABC9A"`` -> ``"ABC9A"``.
    """
    return re.sub(r"\d+$", "", callsign or "").strip()


def lookup_callsign(
    callsign: str, lookup: Optional[Dict[str, Dict[str, str]]] = None
) -> Dict[str, str]:
    """Resolve a dataset callsign against the lookup table.

    The dataset callsigns carry a trailing channel number (e.g. ``"TNQ40"``),
    while the lookup table is keyed on the bare callsign (``"TNQ"``). We try
    the full callsign first, then the prefix with trailing digits stripped.

    Returns ``{"fullName": ..., "name": ...}``. When nothing matches, both
    fields fall back to the callsign prefix so the output is never empty.
    """
    lookup = lookup or {}
    raw = (callsign or "").strip()
    prefix = callsign_prefix(raw)

    for key in (raw.upper(), prefix.upper()):
        if key and key in lookup:
            entry = lookup[key]
            return {
                "fullName": entry.get("fullName") or prefix or raw,
                "name": entry.get("name") or prefix or raw,
            }

    # Unknown callsign: use the prefix for both fields.
    fallback = prefix or raw
    return {"fullName": fallback, "name": fallback}


def build_source_name(channel: Any, frequency_mhz: Any, network_hint: str) -> str:
    """Build the ``source_name`` field.

    The format is ``CH<channel>_<freq*10>_<network_hint>``, e.g.
    ``CH40_6135_Ten``. ``network_hint`` is the ``name`` resolved from the
    lookup table (e.g. ``"Ten"``).
    """
    try:
        freq_tenths = int(round(float(frequency_mhz) * 10))
    except (TypeError, ValueError):
        freq_tenths = 0
    return f"CH{channel}_{freq_tenths}_{network_hint}"


def _to_int_channel(channel: Any) -> Any:
    """Convert a channel value to int when possible, else keep as-is."""
    try:
        return int(str(channel).strip())
    except (TypeError, ValueError):
        return channel


def _to_float(value: Any) -> Optional[float]:
    """Convert a value to float when possible, else ``None``."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def build_area_config(
    data: Dict[str, Dict[str, Any]],
    state: str,
    area: str,
    record_duration_seconds: int = RECORD_DURATION_SECONDS,
    lookup: Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Build the YAML-ready dictionary for one area.

    ``lookup`` is the callsign -> name table produced by
    :func:`load_lookup_table`; it supplies ``FullName`` and ``network_hint``.

    Raises ``KeyError`` if the state or area is not present.
    """
    if state not in data:
        raise KeyError(f"State not found in dataset: {state}")
    if area not in data[state]:
        raise KeyError(f"Area '{area}' not found in state '{state}'")

    site = data[state][area]
    transmissions = site.get("Transmissions", []) or []

    frequencies: List[Dict[str, Any]] = []
    for tx in transmissions:
        channel = tx.get("Channel")
        freq_mhz = _to_float(tx.get("Frequency(MHz)"))
        callsign = tx.get("Callsign", "") or ""

        if freq_mhz is None:
            # Skip entries without a usable frequency.
            continue

        resolved = lookup_callsign(callsign, lookup)
        network_hint = resolved["name"]
        full_name = resolved["fullName"]

        frequencies.append(
            {
                "channel": _to_int_channel(channel),
                "frequency_hz": int(round(freq_mhz * 1_000_000)),
                "frequency_mhz": freq_mhz,
                "bandwidth_mhz": BANDWIDTH_MHZ,
                "standard": STANDARD,
                "Callsign": callsign_prefix(callsign),
                "FullName": full_name,
                "network_hint": network_hint,
                "source_name": build_source_name(channel, freq_mhz, network_hint),
            }
        )

    # Sort by channel number (numeric where possible) for a stable output.
    def _sort_key(item: Dict[str, Any]) -> Any:
        ch = item.get("channel")
        return (0, ch) if isinstance(ch, int) else (1, str(ch))

    frequencies.sort(key=_sort_key)

    return {
        "location": area,
        "country": COUNTRY,
        "region": state,
        "record_duration_seconds": record_duration_seconds,
        "frequencies": frequencies,
    }


# ---------------------------------------------------------------------------
# YAML serialisation
# ---------------------------------------------------------------------------

def _yaml_scalar(value: Any) -> str:
    """Render a scalar value as a YAML-safe string."""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return repr(value) if isinstance(value, float) else str(value)
    text = str(value)
    # Quote strings that could be misinterpreted by YAML parsers.
    # A leading "-", "?", ":" or special indicator needs quoting, as does any
    # string containing structural characters. A hyphen in the middle of a
    # value (e.g. "DVB-T") is safe and is left unquoted, matching the examples.
    needs_quote = (
        text == ""
        or text.strip() != text
        or text[0] in "-?:,[]{}#&*!|>'\"%@`"
        or re.search(r":\s|:\s*$|[\[\]{},#&*!|>'\"%@`]", text) is not None
    )
    if needs_quote:
        escaped = text.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    return text


def to_yaml(config: Dict[str, Any]) -> str:
    """Serialise the area config to YAML matching the example layout.

    The output intentionally mirrors ``Goldcoast_示例.yaml``::

        location: Gold Coast

        country: AU
        region: Gold Coast

        record_duration_seconds: 90

        frequencies:
          - channel: 40
            frequency_hz: 613500000
            ...
    """
    lines: List[str] = []

    lines.append(f"location: {_yaml_scalar(config['location'])}")
    lines.append("")
    lines.append(f"country: {_yaml_scalar(config['country'])}")
    lines.append(f"region: {_yaml_scalar(config['region'])}")
    lines.append("")
    lines.append(f"record_duration_seconds: {_yaml_scalar(config['record_duration_seconds'])}")
    lines.append("")
    lines.append("frequencies:")

    for entry in config["frequencies"]:
        lines.append(f"  - channel: {_yaml_scalar(entry['channel'])}")
        lines.append(f"    frequency_hz: {_yaml_scalar(entry['frequency_hz'])}")
        lines.append(f"    frequency_mhz: {_yaml_scalar(entry['frequency_mhz'])}")
        lines.append(f"    bandwidth_mhz: {_yaml_scalar(entry['bandwidth_mhz'])}")
        lines.append(f"    standard: {_yaml_scalar(entry['standard'])}")
        lines.append(f"    Callsign: {_yaml_scalar(entry['Callsign'])}")
        lines.append(f"    FullName: {_yaml_scalar(entry['FullName'])}")
        lines.append(f"    network_hint: {_yaml_scalar(entry['network_hint'])}")
        lines.append(f"    source_name: {_yaml_scalar(entry['source_name'])}")
        lines.append("")

    # Remove the trailing blank line for a tidy file ending with a newline.
    while lines and lines[-1] == "":
        lines.pop()

    return "\n".join(lines) + "\n"


def safe_filename(name: str) -> str:
    """Turn an area name into a safe file name (without extension)."""
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", name).strip().strip(".")
    return cleaned or "output"


def generate_yaml_file(
    data: Dict[str, Dict[str, Any]],
    state: str,
    area: str,
    output_path: str,
    record_duration_seconds: int = RECORD_DURATION_SECONDS,
    lookup: Optional[Dict[str, Dict[str, str]]] = None,
) -> str:
    """Build the config for ``area`` and write it to ``output_path``.

    Returns the path written.
    """
    config = build_area_config(
        data,
        state,
        area,
        record_duration_seconds=record_duration_seconds,
        lookup=lookup,
    )
    content = to_yaml(config)

    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    with open(output_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(content)

    return output_path


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def run_gui(
    data_file: str, lookup_file: Optional[str] = None
) -> None:
    """Launch the Tkinter GUI."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    # Load the dataset up-front so errors surface immediately.
    try:
        data = load_dataset(data_file)
    except Exception as exc:  # noqa: BLE001 - show any load error to the user
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Failed to load dataset", str(exc))
        root.destroy()
        return

    # Load the callsign -> name lookup table (missing file is tolerated).
    lookup = load_lookup_table(lookup_file) if lookup_file else {}

    root = tk.Tk()
    root.title("DTV YAML Generator")
    root.geometry("680x820")
    root.minsize(600, 720)

    # --- Header -----------------------------------------------------------
    header = ttk.Label(
        root,
        text="Generate a per-area DTV frequency YAML file",
        font=("Segoe UI", 12, "bold"),
    )
    header.pack(padx=16, pady=(16, 4), anchor="w")

    subtitle = ttk.Label(
        root,
        text=f"Data source: {os.path.basename(data_file)}",
        foreground="#666666",
    )
    subtitle.pack(padx=16, pady=(0, 12), anchor="w")

    # --- Form -------------------------------------------------------------
    form = ttk.Frame(root)
    form.pack(fill="x", padx=16)

    ttk.Label(form, text="State:").grid(row=0, column=0, sticky="w", pady=6)
    state_var = tk.StringVar()
    state_combo = ttk.Combobox(
        form, textvariable=state_var, state="readonly", values=list_states(data)
    )
    state_combo.grid(row=0, column=1, sticky="ew", pady=6, padx=(8, 0))

    ttk.Label(form, text="Area Served:").grid(row=1, column=0, sticky="w", pady=6)
    area_var = tk.StringVar()
    area_combo = ttk.Combobox(form, textvariable=area_var, state="readonly", values=[])
    area_combo.grid(row=1, column=1, sticky="ew", pady=6, padx=(8, 0))

    ttk.Label(form, text="Record duration (s):").grid(
        row=2, column=0, sticky="w", pady=6
    )
    duration_var = tk.StringVar(value=str(RECORD_DURATION_SECONDS))
    duration_entry = ttk.Entry(form, textvariable=duration_var, width=12)
    duration_entry.grid(row=2, column=1, sticky="w", pady=6, padx=(8, 0))

    form.columnconfigure(1, weight=1)

    # --- Nearest candidates ----------------------------------------------
    # Straight-line distance does not always match the official broadcast
    # service area, so show the closest sites and let the user confirm.
    candidates_frame = ttk.LabelFrame(
        root, text="Nearest sites (double-click to select)"
    )
    candidates_frame.pack(fill="x", padx=16, pady=(12, 0))

    candidates_list = tk.Listbox(
        candidates_frame, height=5, font=("Consolas", 9), exportselection=False
    )
    candidates_scroll = ttk.Scrollbar(
        candidates_frame, orient="vertical", command=candidates_list.yview
    )
    candidates_list.configure(yscrollcommand=candidates_scroll.set)
    candidates_list.pack(side="left", fill="x", expand=True, padx=(6, 0), pady=6)
    candidates_scroll.pack(side="right", fill="y", pady=6, padx=(0, 6))

    #: Maps a listbox row index -> (state, area).
    candidate_rows: List[Tuple[str, str]] = []

    # --- mySwitch web lookup ---------------------------------------------
    # Opens the official Australian DTV coverage lookup page in the default
    # browser. An optional address pre-fills the mySwitch search box.
    myswitch_frame = ttk.LabelFrame(
        root, text="Official DTV coverage lookup (mySwitch)"
    )
    myswitch_frame.pack(fill="x", padx=16, pady=(12, 0))

    myswitch_row = ttk.Frame(myswitch_frame)
    myswitch_row.pack(fill="x", padx=6, pady=6)

    ttk.Label(myswitch_row, text="Address:").pack(side="left")
    myswitch_address_var = tk.StringVar()
    myswitch_entry = ttk.Entry(myswitch_row, textvariable=myswitch_address_var)
    myswitch_entry.pack(side="left", fill="x", expand=True, padx=(8, 8))

    def on_open_myswitch() -> None:
        address = myswitch_address_var.get().strip()
        if open_myswitch(address):
            if address:
                status_var.set(
                    f"Opened mySwitch in your browser for: {address}"
                )
            else:
                status_var.set("Opened mySwitch in your browser.")
        else:
            messagebox.showerror(
                "Could not open browser",
                "Failed to launch the default web browser.\n"
                f"Please open this page manually:\n{build_myswitch_url(address)}",
            )

    myswitch_btn = ttk.Button(
        myswitch_row, text="Open mySwitch", command=on_open_myswitch
    )
    myswitch_btn.pack(side="right")

    # Pressing Enter in the address box also opens the page.
    myswitch_entry.bind("<Return>", lambda _e: on_open_myswitch())

    myswitch_hint = ttk.Label(
        myswitch_frame,
        text="Enter an address to pre-fill the search, or leave blank to open "
        "the mySwitch home page.",
        foreground="#666666",
    )
    myswitch_hint.pack(fill="x", padx=6, pady=(0, 6))

    # --- Preview ----------------------------------------------------------
    preview_frame = ttk.LabelFrame(root, text="Preview")
    preview_frame.pack(fill="both", expand=True, padx=16, pady=12)

    preview_text = tk.Text(preview_frame, height=8, wrap="none", font=("Consolas", 9))
    preview_scroll = ttk.Scrollbar(
        preview_frame, orient="vertical", command=preview_text.yview
    )
    preview_text.configure(yscrollcommand=preview_scroll.set)
    preview_text.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
    preview_scroll.pack(side="right", fill="y", pady=6, padx=(0, 6))

    # --- Status -----------------------------------------------------------
    status_var = tk.StringVar(value="Select a state and an area.")
    status = ttk.Label(root, textvariable=status_var, foreground="#444444")
    status.pack(fill="x", padx=16)

    # --- Behaviour --------------------------------------------------------
    def on_state_change(_event: Optional[Any] = None) -> None:
        state = state_var.get()
        areas = list_areas(data, state)
        area_combo.configure(values=areas)
        area_var.set("")
        preview_text.delete("1.0", "end")
        if areas:
            status_var.set(f"{len(areas)} area(s) available in {state}.")
        else:
            status_var.set(f"No areas found for {state}.")

    def on_candidate_select(_event: Optional[Any] = None) -> None:
        selection = candidates_list.curselection()
        if not selection:
            return
        index = selection[0]
        if 0 <= index < len(candidate_rows):
            cand_state, cand_area = candidate_rows[index]
            select_record(cand_state, cand_area)

    def on_area_change(_event: Optional[Any] = None) -> None:
        state = state_var.get()
        area = area_var.get()
        if not state or not area:
            return
        try:
            duration = int(duration_var.get())
        except ValueError:
            duration = RECORD_DURATION_SECONDS
        try:
            config = build_area_config(
                data,
                state,
                area,
                record_duration_seconds=duration,
                lookup=lookup,
            )
        except Exception as exc:  # noqa: BLE001
            preview_text.delete("1.0", "end")
            preview_text.insert("1.0", f"Error: {exc}")
            status_var.set("Failed to build preview.")
            return

        preview_text.delete("1.0", "end")
        preview_text.insert("1.0", to_yaml(config))
        status_var.set(
            f"{len(config['frequencies'])} frequency(ies) for {area}, {state}."
        )

    def on_generate() -> None:
        state = state_var.get()
        area = area_var.get()
        if not state or not area:
            messagebox.showwarning(
                "Missing selection", "Please select both a state and an area."
            )
            return

        try:
            duration = int(duration_var.get())
            if duration <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Invalid duration",
                "Record duration must be a positive integer (seconds).",
            )
            return

        default_name = f"{safe_filename(area)}.yaml"
        output_path = filedialog.asksaveasfilename(
            title="Save YAML file",
            defaultextension=".yaml",
            initialfile=default_name,
            filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")],
        )
        if not output_path:
            return

        try:
            generate_yaml_file(
                data,
                state,
                area,
                output_path,
                record_duration_seconds=duration,
                lookup=lookup,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Generation failed", str(exc))
            status_var.set("Generation failed.")
            return

        status_var.set(f"Saved: {output_path}")
        messagebox.showinfo("Success", f"YAML file generated:\n{output_path}")

    state_combo.bind("<<ComboboxSelected>>", on_state_change)
    area_combo.bind("<<ComboboxSelected>>", on_area_change)
    candidates_list.bind("<<ListboxSelect>>", on_candidate_select)
    candidates_list.bind("<Double-Button-1>", on_candidate_select)

    # --- Buttons ----------------------------------------------------------
    buttons = ttk.Frame(root)
    buttons.pack(fill="x", padx=16, pady=(4, 16))

    generate_btn = ttk.Button(buttons, text="Generate YAML", command=on_generate)
    generate_btn.pack(side="right")

    quit_btn = ttk.Button(buttons, text="Quit", command=root.destroy)
    quit_btn.pack(side="right", padx=(0, 8))

    # --- Initial selection ------------------------------------------------
    # Try to locate the user and preselect the nearest transmitter site.
    # Fall back to the first record in the dataset when geolocation fails.
    def select_record(state: str, area: str) -> None:
        state_var.set(state)
        on_state_change()
        area_combo.configure(values=list_areas(data, state))
        area_var.set(area)
        on_area_change()

    def populate_candidates(
        candidates: List[Tuple[str, str, float]]
    ) -> None:
        """Fill the nearest-sites list box and highlight the first entry."""
        candidate_rows.clear()
        candidates_list.delete(0, "end")
        for cand_state, cand_area, distance_km in candidates:
            candidate_rows.append((cand_state, cand_area))
            candidates_list.insert(
                "end", f"{distance_km:6.1f} km  {cand_area}, {cand_state}"
            )
        if candidate_rows:
            candidates_list.selection_set(0)
            candidates_list.activate(0)

    status_var.set("Determining your location...")
    root.update_idletasks()

    location = get_current_location()
    candidates = (
        find_nearest_sites(data, *location, limit=5) if location else []
    )

    if candidates:
        populate_candidates(candidates)
        nearest_state, nearest_area, distance_km = candidates[0]
        select_record(nearest_state, nearest_area)
        status_var.set(
            f"Nearest site: {nearest_area}, {nearest_state} "
            f"(~{distance_km:.0f} km away). "
            "Confirm or pick another from the list above."
        )
    else:
        fallback = first_record(data)
        if fallback is not None:
            select_record(*fallback)
            status_var.set(
                "Location unavailable; showing the first record "
                f"({fallback[1]}, {fallback[0]})."
            )
        else:
            status_var.set("No records found in the dataset.")

    root.mainloop()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _resolve_data_file(explicit: Optional[str]) -> str:
    """Resolve the dataset path, preferring an explicit value."""
    if explicit:
        return explicit
    # Prefer a file next to this script, then the current working directory.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(script_dir, DEFAULT_DATA_FILE)
    if os.path.isfile(candidate):
        return candidate
    return DEFAULT_DATA_FILE


def _resolve_lookup_file(explicit: Optional[str]) -> str:
    """Resolve the callsign lookup-table path, preferring an explicit value.

    Search order: explicit value, then a file next to this script, then the
    current working directory, then the parent directory (the lookup table
    usually lives alongside the dataset at the project root).
    """
    if explicit:
        return explicit
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = (
        os.path.join(script_dir, DEFAULT_LOOKUP_FILE),
        DEFAULT_LOOKUP_FILE,
        os.path.join(script_dir, os.pardir, DEFAULT_LOOKUP_FILE),
    )
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return DEFAULT_LOOKUP_FILE


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Generate per-area DTV frequency YAML files from the "
        "merged broadcast-transmitter dataset."
    )
    parser.add_argument(
        "--data",
        default=None,
        help=f"Path to the merged JSON dataset (default: {DEFAULT_DATA_FILE}).",
    )
    parser.add_argument(
        "--lookup",
        default=None,
        help="Path to the callsign -> name lookup table "
        f"(default: {DEFAULT_LOOKUP_FILE}).",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List all states and areas, then exit.",
    )
    parser.add_argument(
        "--nearest",
        action="store_true",
        help="Detect the current location and print the nearest transmitter "
        "site, then exit.",
    )
    parser.add_argument("--state", help="State to export (e.g. QLD).")
    parser.add_argument("--area", help="Area served to export (e.g. 'Gold Coast').")
    parser.add_argument(
        "--out",
        help="Output YAML path. Defaults to '<area>.yaml' in the current directory.",
    )
    parser.add_argument(
        "--duration",
        type=int,
        default=RECORD_DURATION_SECONDS,
        help=f"record_duration_seconds value (default: {RECORD_DURATION_SECONDS}).",
    )

    args = parser.parse_args(argv)
    data_file = _resolve_data_file(args.data)
    lookup_file = _resolve_lookup_file(args.lookup)

    try:
        data = load_dataset(data_file)
    except Exception as exc:  # noqa: BLE001
        print(f"Error loading dataset: {exc}", file=sys.stderr)
        return 1

    lookup = load_lookup_table(lookup_file)

    if args.list:
        for state in list_states(data):
            areas = list_areas(data, state)
            print(f"{state} ({len(areas)} areas)")
            for area in areas:
                print(f"    {area}")
        return 0

    if args.nearest:
        location = get_current_location()
        if location is None:
            print("Could not determine current location.", file=sys.stderr)
            fallback = first_record(data)
            if fallback is not None:
                print(f"First record: {fallback[1]}, {fallback[0]}")
            return 1
        nearest = find_nearest_site(data, *location)
        if nearest is None:
            print("No transmitter sites with coordinates found.", file=sys.stderr)
            return 1
        nearest_state, nearest_area, distance_km = nearest
        print(f"Location: {location[0]:.4f}, {location[1]:.4f}")
        print(
            f"Nearest site: {nearest_area}, {nearest_state} "
            f"(~{distance_km:.0f} km away)"
        )
        return 0

    # If no state/area given, launch the GUI.
    if not args.state or not args.area:
        run_gui(data_file, lookup_file)
        return 0

    out_path = args.out or f"{safe_filename(args.area)}.yaml"
    try:
        written = generate_yaml_file(
            data,
            args.state,
            args.area,
            out_path,
            record_duration_seconds=args.duration,
            lookup=lookup,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"Error generating YAML: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {written}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
