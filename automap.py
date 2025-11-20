"""
Script to calculate driving distances and travel times between a single origin address
and a list of destination addresses using the openrouteservice API.

A valid openrouteservice API key must be provided in the environment variable
`OPENROUTESERVICE_API_KEY` before running the script.

To set the environment variable OPENROUTESERVICE_API_KEY:
  - macOS / Linux (bash): export OPENROUTESERVICE_API_KEY="your_api_key"
  - Windows (PowerShell): setx OPENROUTESERVICE_API_KEY "your_api_key"
"""

from __future__ import annotations

import os
import time
from typing import Optional, Tuple
from numbers import Number

import pandas as pd
import requests

# Main configuration section, easy to adapt.
INPUT_FILE = "addresses.xlsx"
OUTPUT_FILE = "addresses_with_distances.xlsx"
ADDRESS_COLUMN = "Address"
POSTAL_CODE_COLUMN = "PostalCode"
CITY_COLUMN = "City"
ORIGIN_ADDRESS = "Origin address - ZIP City"
HEADER_ROW = 0  # Row index (0-based) containing the column headers in Excel.
COUNTRY_HINT = "Switzerland"
API_KEY = os.getenv("OPENROUTESERVICE_API_KEY")
GEOCODE_URL = "https://api.openrouteservice.org/geocode/search"
DIRECTIONS_URL = "https://api.openrouteservice.org/v2/directions/driving-car"
SLEEP_BETWEEN_CALLS = 0.2  # Small delay to avoid hitting the API too hard.


def _normalize_field(value: object) -> str:
    """Cleans up a value coming from Excel (keeps integer postal codes without .0)."""

    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if pd.isna(value):
        return ""
    if isinstance(value, Number):
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            return str(value).strip()
        if numeric_value.is_integer():
            return str(int(numeric_value))
        return str(numeric_value)
    return str(value).strip()


def build_full_address(street: object, postal_code: object, city: object) -> str:
    """Assemble street, postal code, city and country to improve geocoding."""

    parts = []
    street_part = _normalize_field(street)
    if street_part:
        parts.append(street_part)

    locality_parts = []
    postal_part = _normalize_field(postal_code)
    city_part = _normalize_field(city)
    if postal_part:
        locality_parts.append(postal_part)
    if city_part:
        locality_parts.append(city_part)
    if locality_parts:
        parts.append(" ".join(locality_parts))

    if COUNTRY_HINT:
        parts.append(COUNTRY_HINT)

    return ", ".join(parts)


def geocode_address(address: str, api_key: str) -> Optional[Tuple[float, float]]:
    """Return (lon, lat) for a given address using the geocoding API."""

    params = {
        "api_key": api_key,
        "text": address,
        "size": 1,
        "boundary.country": "CH",
    }

    try:
        response = requests.get(GEOCODE_URL, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
    except requests.RequestException:
        return None

    features = data.get("features")
    if not features:
        return None

    geometry = features[0].get("geometry") or {}
    coordinates = geometry.get("coordinates")
    if not coordinates or len(coordinates) < 2:
        return None

    lon, lat = coordinates[0], coordinates[1]
    return lon, lat


def get_distance_and_duration(origin: str, destination: str, api_key: str) -> Tuple[Optional[float], Optional[float]]:
    """Call the openrouteservice APIs to retrieve distance (km) and duration (minutes).

    Returns (distance_km, duration_minutes) or (None, None) in case of error.
    Numerical values are rounded to 2 decimals for distance and 1 decimal for duration.
    """

    origin_coords = geocode_address(origin, api_key)
    destination_coords = geocode_address(destination, api_key)
    if not origin_coords or not destination_coords:
        return None, None

    payload = {"coordinates": [origin_coords, destination_coords]}
    headers = {"Authorization": api_key, "Content-Type": "application/json"}

    try:
        response = requests.post(DIRECTIONS_URL, json=payload, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
    except requests.RequestException:
        return None, None

    routes = data.get("routes")
    if not routes:
        return None, None

    summary = routes[0].get("summary") or {}
    distance_m = summary.get("distance")
    duration_s = summary.get("duration")
    if distance_m is None or duration_s is None:
        return None, None

    distance_km = round(distance_m / 1000, 2)
    duration_minutes = round(duration_s / 60, 1)
    return distance_km, duration_minutes


def main() -> None:
    """Load the Excel file, perform the API calls and save the result."""
    if not API_KEY:
        print("Error: environment variable OPENROUTESERVICE_API_KEY is missing.")
        return

    try:
        df = pd.read_excel(INPUT_FILE, header=HEADER_ROW)
    except FileNotFoundError:
        print(f"Error: input file '{INPUT_FILE}' not found.")
        return
    except Exception as exc:  # pragma: no cover - general safeguard
        print(f"Error while reading the Excel file: {exc}")
        return

    required_columns = [ADDRESS_COLUMN, POSTAL_CODE_COLUMN, CITY_COLUMN]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        available = ", ".join(map(str, df.columns)) or "(no columns detected)"
        print(
            "Error: missing columns: "
            f"{', '.join(missing_columns)}. Available columns: {available}."
        )
        return

    df["Distance_km"] = ""
    df["Duration_minutes"] = ""

    print(f"Processing {len(df)} rows...")
    for idx, row in df.iterrows():
        if idx % 10 == 0:
            print(f"Row {idx + 1}/{len(df)}")

        address = row.get(ADDRESS_COLUMN, "")
        if pd.isna(address) or not str(address).strip():
            df.at[idx, "Distance_km"] = "NO_ADDRESS"
            df.at[idx, "Duration_minutes"] = "NO_ADDRESS"
            continue

        postal_code = row.get(POSTAL_CODE_COLUMN, "")
        city = row.get(CITY_COLUMN, "")
        full_destination = build_full_address(address, postal_code, city)
        if not full_destination:
            df.at[idx, "Distance_km"] = "NO_ADDRESS"
            df.at[idx, "Duration_minutes"] = "NO_ADDRESS"
            continue

        distance_km, duration_minutes = get_distance_and_duration(
            ORIGIN_ADDRESS,
            full_destination,
            API_KEY,
        )

        if distance_km is None or duration_minutes is None:
            df.at[idx, "Distance_km"] = "API_ERROR"
            df.at[idx, "Duration_minutes"] = "API_ERROR"
        else:
            df.at[idx, "Distance_km"] = distance_km
            df.at[idx, "Duration_minutes"] = duration_minutes

        time.sleep(SLEEP_BETWEEN_CALLS)

    try:
        df.to_excel(OUTPUT_FILE, index=False)
    except Exception as exc:  # pragma: no cover - general safeguard
        print(f"Error while writing the output file: {exc}")
        return

    print(f"Output file with distances saved as '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    main()
