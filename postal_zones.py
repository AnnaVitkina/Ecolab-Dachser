"""Build postal-code zone TXT export from cleaned Dachser zone headers."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ZONE_COUNTRY = "IE"
EXCLUDED_PLACEHOLDER = "-"
ZONE_E_IE_POSTAL_CODES = "F92, F93, F94"
ZONE_E_GB_POSTAL_CODES = "BT"


@dataclass(frozen=True)
class PostalZone:
    name: str
    country: str
    postal_codes: str
    excluded: str = EXCLUDED_PLACEHOLDER


def _zone_letter(zone_header: str) -> str:
    first_line = zone_header.split("\n", 1)[0].strip()
    match = re.search(r"Zone\s+([A-E])", first_line, re.IGNORECASE)
    return match.group(1).upper() if match else first_line


def _postal_codes_from_header(zone_header: str) -> str:
    lines = [line.strip() for line in zone_header.splitlines() if line.strip()]
    if len(lines) <= 1:
        return ""
    return ", ".join(lines[1:])


def _zones_for_header(zone_header: str) -> list[PostalZone]:
    letter = _zone_letter(zone_header)
    if letter == "E":
        return [
            PostalZone(
                name="IE Zone E",
                country="IE",
                postal_codes=ZONE_E_IE_POSTAL_CODES,
            ),
            PostalZone(
                name="GB Zone E",
                country="GB",
                postal_codes=ZONE_E_GB_POSTAL_CODES,
            ),
        ]
    return [
        PostalZone(
            name=f"IE Zone {letter}",
            country=ZONE_COUNTRY,
            postal_codes=_postal_codes_from_header(zone_header),
        )
    ]


def postal_zones_from_fields(fields: dict[str, Any]) -> list[PostalZone]:
    main_costs = fields.get("MainCosts") or []
    if not main_costs:
        raise ValueError("MainCosts is missing or empty.")

    header_row = main_costs[0]
    zones: list[PostalZone] = []
    for key in sorted(header_row):
        if not str(key).startswith("Zone"):
            continue
        zone_header = str(header_row.get(key, "")).strip()
        if not zone_header:
            continue
        zones.extend(_zones_for_header(zone_header))
    if not zones:
        raise ValueError("No zone columns found in MainCosts header row.")
    return zones


def format_postal_zone_txt(zone: PostalZone) -> str:
    return (
        f"Name\n{zone.name}\n"
        f"Country\n{zone.country}\n"
        f"Postal Code\n{zone.postal_codes}\n"
        f"Excluded\n{zone.excluded}\n"
    )


def write_postal_zones_txt(zones: list[PostalZone], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        for index, zone in enumerate(zones):
            handle.write(format_postal_zone_txt(zone))
            if index < len(zones) - 1:
                handle.write("\n")
    return output_path


def postal_zones_txt_path_for_matrix(matrix_path: Path) -> Path:
    return matrix_path.with_suffix(".txt")
