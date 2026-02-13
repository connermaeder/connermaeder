#!/usr/bin/env python3
"""
Draw and optionally optimize a route from a CSV file on an interactive map.

Default behavior:
- Reads route.csv from the current folder
- Uses geocode_cache.json when available
- Draws route_map.html with OpenStreetMap + Leaflet
- Tries OR-Tools in auto mode, falls back to heuristic if OR-Tools is unavailable
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple


DEFAULT_INPUT = "route.csv"
DEFAULT_CACHE = "geocode_cache.json"
DEFAULT_OUTPUT_HTML = "route_map.html"
DEFAULT_OUTPUT_MISSING = "route_missing.csv"
DEFAULT_USER_AGENT = "route-planner/1.0 (local-script)"


@dataclass
class RouteRow:
    address_raw: str
    stop_nr: Optional[int]
    hint: str


@dataclass
class ResolvedPoint:
    address_raw: str
    stop_nr: Optional[int]
    hint: str
    lat: float
    lon: float


def normalize_whitespace(value: str) -> str:
    return " ".join((value or "").replace("\u00a0", " ").split()).strip()


def fix_common_mojibake(value: str) -> str:
    replacements = {
        "Ã¼": "u",
        "Ã¶": "o",
        "Ã¤": "a",
        "Ã": "U",
        "Ã": "O",
        "Ã": "A",
        "Ã": "ss",
        "â": "-",
        "â": "-",
    }
    result = value
    for wrong, right in replacements.items():
        result = result.replace(wrong, right)
    return result


def normalize_key(value: str) -> str:
    text = normalize_whitespace(value)
    text = fix_common_mojibake(text)
    text = text.casefold()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^0-9a-z, ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace(" ,", ",")
    return text


def parse_int(value: str) -> Optional[int]:
    text = normalize_whitespace(value)
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def read_csv_rows(path: Path) -> Tuple[List[Dict[str, str]], List[str]]:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                sample = handle.read(8192)
                handle.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=";,")
                except csv.Error:
                    dialect = csv.excel

                reader = csv.DictReader(handle, dialect=dialect)
                if not reader.fieldnames:
                    continue
                rows = [dict(row) for row in reader]
                headers = [str(h) for h in reader.fieldnames]
                return rows, headers
        except Exception as exc:
            last_error = exc

    if last_error is not None:
        raise RuntimeError(f"Failed to read CSV: {last_error}") from last_error
    raise RuntimeError("Failed to read CSV.")


def detect_column(headers: Sequence[str], requested: Optional[str], candidates: Sequence[str]) -> Optional[str]:
    if requested:
        if requested in headers:
            return requested
        raise ValueError(f"Column '{requested}' not found. Available: {headers}")

    normalized = {normalize_whitespace(name).casefold(): name for name in headers}
    for candidate in candidates:
        found = normalized.get(candidate.casefold())
        if found:
            return found

    for name in headers:
        lowered = normalize_whitespace(name).casefold()
        if any(candidate.casefold() in lowered for candidate in candidates):
            return name

    return None


def read_route_rows(path: Path, address_column: Optional[str]) -> List[RouteRow]:
    rows, headers = read_csv_rows(path)
    if not rows:
        return []

    address_col = detect_column(headers, address_column, ("Adresse_raw", "Adresse", "Address", "Addr"))
    if not address_col:
        address_col = headers[0]

    stop_col = detect_column(headers, None, ("StopNr", "Stop", "Nr"))
    hint_col = detect_column(headers, None, ("Hinweis", "Note", "Info"))

    parsed: List[RouteRow] = []
    for row in rows:
        address = normalize_whitespace(row.get(address_col, ""))
        if not address:
            continue
        stop_nr = parse_int(row.get(stop_col, "")) if stop_col else None
        hint = normalize_whitespace(row.get(hint_col, "")) if hint_col else ""
        parsed.append(RouteRow(address_raw=address, stop_nr=stop_nr, hint=hint))

    if any(item.stop_nr is not None for item in parsed):
        parsed.sort(key=lambda item: (item.stop_nr is None, item.stop_nr if item.stop_nr is not None else 10**9))

    return parsed


def is_start_row(row: RouteRow) -> bool:
    hint = row.hint.casefold()
    return row.stop_nr == 0 or hint == "start" or "start" in hint


def move_start_to_front(rows: Sequence[RouteRow]) -> List[RouteRow]:
    if not rows:
        return []

    start_index = 0
    for idx, row in enumerate(rows):
        if is_start_row(row):
            start_index = idx
            break

    if start_index == 0:
        return list(rows)

    reordered = [rows[start_index]]
    reordered.extend(rows[:start_index])
    reordered.extend(rows[start_index + 1 :])
    return reordered


def load_cache(path: Path) -> Dict[str, Dict[str, float]]:
    if not path.exists():
        return {}

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}

    if not isinstance(payload, dict):
        return {}

    cache: Dict[str, Dict[str, float]] = {}
    for key, value in payload.items():
        if not isinstance(value, dict):
            continue
        lat = value.get("lat")
        lon = value.get("lon")
        if isinstance(lat, (int, float)) and isinstance(lon, (int, float)):
            cache[normalize_key(str(key))] = {"lat": float(lat), "lon": float(lon)}
    return cache


def save_cache(path: Path, cache: Dict[str, Dict[str, float]]) -> None:
    path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


class NominatimGeocoder:
    def __init__(
        self,
        user_agent: str,
        sleep_seconds: float = 1.1,
        countrycodes: str = "ch",
    ) -> None:
        self.user_agent = user_agent
        self.sleep_seconds = max(0.0, sleep_seconds)
        self.countrycodes = countrycodes
        self._last_call = 0.0

    def geocode(self, query: str, retries: int = 3) -> Optional[Tuple[float, float]]:
        for attempt in range(retries):
            wait_seconds = self.sleep_seconds - (time.time() - self._last_call)
            if wait_seconds > 0:
                time.sleep(wait_seconds)
            self._last_call = time.time()

            try:
                params = {
                    "q": query,
                    "format": "jsonv2",
                    "limit": 1,
                }
                if self.countrycodes:
                    params["countrycodes"] = self.countrycodes

                url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(params)
                request = urllib.request.Request(
                    url,
                    headers={
                        "User-Agent": self.user_agent,
                        "Accept": "application/json",
                    },
                )
                with urllib.request.urlopen(request, timeout=20) as response:
                    payload = response.read().decode("utf-8")
                data = json.loads(payload)

                if not data:
                    return None

                lat = float(data[0]["lat"])
                lon = float(data[0]["lon"])
                return lat, lon
            except (urllib.error.URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError):
                if attempt == retries - 1:
                    return None
                time.sleep(1.5 * (attempt + 1))

        return None


def get_coordinates(
    address: str,
    cache: Dict[str, Dict[str, float]],
    geocoder: NominatimGeocoder,
) -> Optional[Tuple[float, float]]:
    key = normalize_key(address)
    cached = cache.get(key)
    if cached:
        return float(cached["lat"]), float(cached["lon"])

    coords = geocoder.geocode(address)
    if coords:
        cache[key] = {"lat": float(coords[0]), "lon": float(coords[1])}
    return coords


def haversine_km(a: Tuple[float, float], b: Tuple[float, float]) -> float:
    lat1, lon1 = a
    lat2, lon2 = b
    radius_km = 6371.0088

    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)

    x = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    return 2 * radius_km * math.asin(math.sqrt(x))


def route_length_km(coords: Sequence[Tuple[float, float]], order: Sequence[int]) -> float:
    if len(order) < 2:
        return 0.0
    total = 0.0
    for i in range(len(order) - 1):
        total += haversine_km(coords[order[i]], coords[order[i + 1]])
    return total


def nearest_neighbor_open(coords: Sequence[Tuple[float, float]]) -> List[int]:
    if len(coords) <= 2:
        return list(range(len(coords)))

    remaining = set(range(1, len(coords)))
    order = [0]

    while remaining:
        current = order[-1]
        nxt = min(remaining, key=lambda idx: haversine_km(coords[current], coords[idx]))
        order.append(nxt)
        remaining.remove(nxt)

    return order


def two_opt_open(coords: Sequence[Tuple[float, float]], order: List[int], max_passes: int = 25) -> List[int]:
    if len(order) < 4:
        return order

    candidate = list(order)
    n = len(candidate)

    for _ in range(max_passes):
        improved = False
        for i in range(1, n - 2):
            for j in range(i + 1, n - 1):
                a = candidate[i - 1]
                b = candidate[i]
                c = candidate[j]
                d = candidate[j + 1]

                old_len = haversine_km(coords[a], coords[b]) + haversine_km(coords[c], coords[d])
                new_len = haversine_km(coords[a], coords[c]) + haversine_km(coords[b], coords[d])

                if new_len + 1e-9 < old_len:
                    candidate[i : j + 1] = reversed(candidate[i : j + 1])
                    improved = True
        if not improved:
            break

    return candidate


def solve_with_heuristic(coords: Sequence[Tuple[float, float]]) -> List[int]:
    seed = nearest_neighbor_open(coords)
    return two_opt_open(coords, seed)


def solve_with_ortools(coords: Sequence[Tuple[float, float]], time_limit_sec: int) -> Tuple[List[int], Optional[str]]:
    try:
        from ortools.constraint_solver import pywrapcp, routing_enums_pb2
    except ImportError:
        return [], "OR-Tools is not installed. Install with: pip install ortools"

    n = len(coords)
    if n <= 2:
        return list(range(n)), None

    dummy_end = n
    size = n + 1

    matrix: List[List[int]] = [[0] * size for _ in range(size)]
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            matrix[i][j] = int(round(haversine_km(coords[i], coords[j]) * 1000.0))
        matrix[i][dummy_end] = 0

    manager = pywrapcp.RoutingIndexManager(size, 1, [0], [dummy_end])
    routing = pywrapcp.RoutingModel(manager)

    def distance_callback(from_index: int, to_index: int) -> int:
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return matrix[from_node][to_node]

    transit = routing.RegisterTransitCallback(distance_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit)

    search = pywrapcp.DefaultRoutingSearchParameters()
    search.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
    search.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    search.time_limit.FromSeconds(max(1, int(time_limit_sec)))

    solution = routing.SolveWithParameters(search)
    if solution is None:
        return [], "OR-Tools could not find a route solution."

    order: List[int] = []
    index = routing.Start(0)
    while not routing.IsEnd(index):
        node = manager.IndexToNode(index)
        if node < n:
            order.append(node)
        index = solution.Value(routing.NextVar(index))

    if not order or order[0] != 0:
        return [], "OR-Tools returned an invalid route."

    return order, None


def choose_solver(solver: str) -> str:
    if solver != "auto":
        return solver

    try:
        import ortools  # type: ignore  # noqa: F401

        return "ortools"
    except Exception:
        return "heuristic"


def write_missing_csv(path: Path, unresolved: Sequence[RouteRow]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["StopNr", "Adresse_raw", "Hinweis"])
        writer.writeheader()
        for row in unresolved:
            writer.writerow(
                {
                    "StopNr": "" if row.stop_nr is None else row.stop_nr,
                    "Adresse_raw": row.address_raw,
                    "Hinweis": row.hint,
                }
            )


def write_route_map(path: Path, points: Sequence[ResolvedPoint], order: Sequence[int], title: str) -> None:
    ordered_points = [points[i] for i in order]
    if not ordered_points:
        raise ValueError("No points available for map rendering.")

    center_lat = sum(point.lat for point in ordered_points) / len(ordered_points)
    center_lon = sum(point.lon for point in ordered_points) / len(ordered_points)

    js_points: List[Dict[str, object]] = []
    for idx, point in enumerate(ordered_points, start=1):
        stop_label = "" if point.stop_nr is None else str(point.stop_nr)
        hint_suffix = f" ({point.hint})" if point.hint else ""
        popup = html.escape(f"#{idx} | StopNr={stop_label} | {point.address_raw}{hint_suffix}")
        js_points.append(
            {
                "seq": idx,
                "lat": point.lat,
                "lon": point.lon,
                "address": point.address_raw,
                "hint": point.hint,
                "popup": popup,
                "is_start": idx == 1,
            }
        )

    points_json = json.dumps(js_points, ensure_ascii=False)
    title_html = html.escape(title)

    document = f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">
  <title>{title_html}</title>
  <link rel=\"stylesheet\" href=\"https://unpkg.com/leaflet@1.9.4/dist/leaflet.css\" crossorigin=\"\"/>
  <style>
    html, body {{ height: 100%; margin: 0; font-family: Segoe UI, sans-serif; }}
    #map {{ height: 100%; width: 100%; }}
    .legend {{
      position: absolute;
      z-index: 1000;
      top: 12px;
      left: 12px;
      background: rgba(255,255,255,0.95);
      padding: 10px 12px;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.15);
      line-height: 1.4;
      font-size: 13px;
    }}
    .badge {{ display: inline-block; min-width: 18px; text-align: center; border-radius: 10px; color: white; padding: 2px 6px; }}
    .start {{ background: #dc2626; }}
    .stop {{ background: #2563eb; }}
  </style>
</head>
<body>
  <div class=\"legend\">
    <div><strong>{title_html}</strong></div>
    <div><span class=\"badge start\">S</span> Start</div>
    <div><span class=\"badge stop\">#</span> Stop</div>
  </div>
  <div id=\"map\"></div>

  <script src=\"https://unpkg.com/leaflet@1.9.4/dist/leaflet.js\" crossorigin=\"\"></script>
  <script>
    const points = {points_json};

    const map = L.map('map').setView([{center_lat:.6f}, {center_lon:.6f}], 13);
    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
      attribution: '&copy; OpenStreetMap contributors',
      maxZoom: 19
    }}).addTo(map);

    const latlngs = points.map(p => [p.lat, p.lon]);
    const line = L.polyline(latlngs, {{ color: '#2563eb', weight: 4, opacity: 0.85 }}).addTo(map);

    points.forEach((point) => {{
      if (point.is_start) {{
        L.circleMarker([point.lat, point.lon], {{
          radius: 8,
          color: '#991b1b',
          fillColor: '#dc2626',
          fillOpacity: 1,
          weight: 2
        }}).addTo(map).bindPopup(point.popup);
      }} else {{
        L.marker([point.lat, point.lon]).addTo(map).bindPopup(point.popup);
      }}
    }});

    if (latlngs.length > 1) {{
      map.fitBounds(line.getBounds(), {{ padding: [24, 24] }});
    }}
  </script>
</body>
</html>
"""

    path.write_text(document, encoding="utf-8")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Draw a route from route.csv and export an interactive HTML map.")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Input CSV file (default: route.csv)")
    parser.add_argument("--address-column", default=None, help="Address column name. Auto-detected when omitted.")
    parser.add_argument("--cache", default=DEFAULT_CACHE, help="Geocode cache file (default: geocode_cache.json)")
    parser.add_argument("--output-html", default=DEFAULT_OUTPUT_HTML, help="Output HTML map file")
    parser.add_argument("--missing-output", default=DEFAULT_OUTPUT_MISSING, help="Output CSV for unresolved addresses")
    parser.add_argument(
        "--solver",
        choices=("auto", "csv", "heuristic", "ortools"),
        default="auto",
        help="Route order mode: auto | csv | heuristic | ortools",
    )
    parser.add_argument("--ortools-time-limit", type=int, default=20, help="Time limit in seconds for OR-Tools")
    parser.add_argument("--user-agent", default=DEFAULT_USER_AGENT, help="User-Agent for Nominatim requests")
    parser.add_argument("--countrycodes", default="ch", help="Nominatim country filter, e.g. 'ch'")
    parser.add_argument("--sleep-seconds", type=float, default=1.1, help="Delay between geocoding requests")
    parser.add_argument("--title", default="Route map", help="Map title")
    parser.add_argument("--open", action="store_true", help="Open generated map in browser")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path = Path(args.input).expanduser()
    cache_path = Path(args.cache).expanduser()
    output_html = Path(args.output_html).expanduser()
    missing_output = Path(args.missing_output).expanduser()

    if not input_path.exists():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 2

    try:
        route_rows = read_route_rows(input_path, args.address_column)
    except Exception as exc:
        print(f"Error while reading input: {exc}", file=sys.stderr)
        return 2

    if not route_rows:
        print("No addresses found in input file.", file=sys.stderr)
        return 2

    route_rows = move_start_to_front(route_rows)

    geocoder = NominatimGeocoder(
        user_agent=args.user_agent,
        sleep_seconds=args.sleep_seconds,
        countrycodes=args.countrycodes,
    )
    cache = load_cache(cache_path)

    resolved: List[ResolvedPoint] = []
    unresolved: List[RouteRow] = []

    total = len(route_rows)
    for index, row in enumerate(route_rows, start=1):
        coords = get_coordinates(row.address_raw, cache, geocoder)
        if not coords:
            unresolved.append(row)
            print(f"[{index}/{total}] Missing: {row.address_raw}")
            continue

        resolved.append(
            ResolvedPoint(
                address_raw=row.address_raw,
                stop_nr=row.stop_nr,
                hint=row.hint,
                lat=coords[0],
                lon=coords[1],
            )
        )
        print(f"[{index}/{total}] OK: {row.address_raw}")

    save_cache(cache_path, cache)

    if unresolved:
        write_missing_csv(missing_output, unresolved)

    if not resolved:
        print("No geocoded points available. Aborting.", file=sys.stderr)
        return 3

    if len(resolved) == 1:
        order = [0]
        solver_used = "csv"
    else:
        solver_used = choose_solver(args.solver)
        coords = [(point.lat, point.lon) for point in resolved]

        if solver_used == "csv":
            order = list(range(len(resolved)))
        elif solver_used == "heuristic":
            order = solve_with_heuristic(coords)
        elif solver_used == "ortools":
            order, err = solve_with_ortools(coords, args.ortools_time_limit)
            if err:
                print(err, file=sys.stderr)
                if args.solver == "auto":
                    solver_used = "heuristic"
                    order = solve_with_heuristic(coords)
                else:
                    return 4
        else:
            print(f"Unknown solver: {solver_used}", file=sys.stderr)
            return 2

    if len(order) != len(resolved):
        print("Solver returned invalid route size.", file=sys.stderr)
        return 4

    write_route_map(output_html, resolved, order, args.title)

    ordered_coords = [(resolved[i].lat, resolved[i].lon) for i in order]
    route_km = route_length_km(ordered_coords, list(range(len(ordered_coords))))

    print("")
    print(f"Map saved: {output_html}")
    print(f"Solver used: {solver_used}")
    print(f"Stops on map: {len(resolved)}")
    print(f"Unresolved stops: {len(unresolved)}")
    print(f"Approx. route length (straight lines): {route_km:.2f} km")
    if unresolved:
        print(f"Missing list: {missing_output}")

    if args.open:
        webbrowser.open(output_html.resolve().as_uri())

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
