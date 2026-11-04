"""FSD home-signal mappings used by RTIS J/K events."""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from io import BytesIO
import json
import math
import re
import csv

from openpyxl import load_workbook
from sqlalchemy.exc import IntegrityError
from sqlmodel import Field, Session, SQLModel, select


class RtisHomeModel(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    filename: str
    digest: str = Field(unique=True)
    uploaded_at: datetime = Field(default_factory=datetime.utcnow)
    signal_count: int
    station_count: int
    mapped_station_count: int
    mapping_json: str


# A source workbook contained a repeated longitude for the UP Home series from
# TIT through DHN.  Those points form a false straight line and must not be
# shown in the active model, even when an older model record is selected.
_BAD_UP_HOME_LONGITUDE = 88.13194333333334
_BAD_UP_HOME_LATITUDE_MIN = 23.288257 - 0.00001
_BAD_UP_HOME_LATITUDE_MAX = 23.789305 + 0.00001
_NEARBY_DUPLICATE_HOME_METERS = 100.0


def _is_known_bad_up_home(home, group):
    try:
        longitude = float(home.get('longitude'))
        latitude = float(home.get('latitude'))
    except (TypeError, ValueError):
        return False
    return (group.get('event') == 'J'
            and str(home.get('line') or '').strip().upper().startswith('UP')
            and abs(longitude - _BAD_UP_HOME_LONGITUDE) <= 0.000001
            and _BAD_UP_HOME_LATITUDE_MIN <= latitude <= _BAD_UP_HOME_LATITUDE_MAX)


def _remove_known_bad_home_series(mapping):
    """Remove the false TIT-to-DHN UP Home series from an existing model."""
    changed = False
    cleaned = {}
    for key, group in mapping.items():
        homes = [home for home in group.get('homes', []) if not _is_known_bad_up_home(home, group)]
        if len(homes) != len(group.get('homes', [])):
            changed = True
        if not homes:
            changed = True
            continue
        if len(homes) != len(group.get('homes', [])):
            group = dict(group)
            group['homes'] = homes
        cleaned[key] = group
    return cleaned, changed


def _remove_nearby_duplicate_homes(mapping):
    """Keep the first (oldest) nearby signal for a station and line."""
    changed = False
    for group in mapping.values():
        kept = []
        for home in group.get('homes', []):
            line = str(home.get('line') or '').strip().upper()
            signal_type = str(home.get('type') or '').strip().upper()
            duplicate = False
            for previous in kept:
                if line != str(previous.get('line') or '').strip().upper():
                    continue
                if signal_type != str(previous.get('type') or '').strip().upper():
                    continue
                try:
                    distance = _haversine_m(float(home['latitude']), float(home['longitude']),
                                            float(previous['latitude']), float(previous['longitude']))
                except (KeyError, TypeError, ValueError):
                    continue
                if distance <= _NEARBY_DUPLICATE_HOME_METERS:
                    duplicate = True
                    changed = True
                    break
            if not duplicate:
                kept.append(home)
        group['homes'] = kept
    return mapping, changed


def _number(value, label):
    try:
        result = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{label} must be a number.") from None
    if not math.isfinite(result):
        raise ValueError(f"{label} must be a finite number.")
    return result


def _parse_coordinate(value, maximum, label):
    raw = _number(value, label)
    absolute = abs(raw)
    if absolute <= maximum:
        return raw
    degrees = math.floor(absolute / 100)
    minutes = absolute - (degrees * 100)
    if degrees > maximum or minutes >= 60:
        raise ValueError(f"{label} must be decimal degrees or DDMM.mmm format.")
    converted = math.copysign(degrees + minutes / 60, raw)
    if abs(converted) > maximum:
        raise ValueError(f"{label} is out of range.")
    return converted


def _parse_route_home_model_rows(rows, source_name):
    """Parse route CSV/XLSX rows with station/name/railway coordinates."""
    rows = [list(row) for row in rows]
    direction_token = ""
    for row in rows[:3]:
        if len(row) > 3 and str(row[3] or "").strip().upper() in ("U", "D"):
            direction_token = str(row[3]).strip().upper()
            break
    if not direction_token:
        name_token = str(source_name or "").upper()
        direction_token = "U" if "UP" in name_token else "D" if "DN" in name_token or "DOWN" in name_token else ""
    if direction_token not in ("U", "D"):
        raise ValueError("Route file must include U or D in column 4 of its first row.")
    direction = "UP MAIN LINE" if direction_token == "U" else "DN MAIN LINE"
    mapping = {}
    for row_number, row in enumerate(rows, 1):
        if len(row) < 7:
            continue
        station = str(row[1] or "").strip().upper()
        signal_name = str(row[2] or "").strip().upper()
        if not station or not signal_name:
            continue
        signal_type = signal_name[len(station):].strip(" -") if signal_name.startswith(station) else signal_name
        signal_type = {"HOME": "Home", "I/HOME": "I/Home", "INT HOME": "INT Home"}.get(signal_type, signal_type)
        latitude = _parse_coordinate(row[5], 90, f"{source_name} row {row_number} Latitude")
        longitude = _parse_coordinate(row[6], 180, f"{source_name} row {row_number} Longitude")
        event = "J" if direction_token == "U" else "K"
        key = f"{station}|{event}"
        mapping.setdefault(key, {"station": station, "event": event,
                                 "label": "UP HOME" if event == "J" else "DOWN HOME", "homes": []})["homes"].append({
            "line": direction, "type": signal_type, "latitude": latitude, "longitude": longitude,
        })
    if not mapping:
        raise ValueError("No signal rows found in the route file.")
    return mapping, sum(len(value["homes"]) for value in mapping.values())


def parse_home_model(content: bytes, filename: str = ""):
    if len(content) > 25 * 1024 * 1024:
        raise ValueError("FSD file must be 25 MB or smaller.")
    if filename.lower().endswith(".csv"):
        try:
            rows = list(csv.reader(content.decode("utf-8-sig").splitlines()))
        except (UnicodeDecodeError, csv.Error):
            raise ValueError("Cannot read FSD CSV. Upload a UTF-8 CSV file.") from None
        return _parse_route_home_model_rows(rows, filename)
    try:
        book = load_workbook(BytesIO(content), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError("Cannot read FSD file. Upload a valid .xlsx workbook.") from exc
    mapping = {}
    route_sheets = []
    try:
        for sheet in book:
            sheet_rows = list(sheet.iter_rows(values_only=True))
            if not sheet_rows:
                continue
            header, *data_rows = sheet_rows
            labels = [' '.join(str(value or '').upper().split()) for value in header]
            direction_name = next((name for name in ('DIRN', 'STATION DIRN', 'DIRECTION') if name in labels), None)
            required = ('STATION', direction_name, 'TYPE', 'LATITUDE', 'LONGITUDE')
            if not direction_name or not all(name in labels for name in required):
                route_sheets.append((sheet.title, sheet_rows))
                continue
            indexes = {name: labels.index(name) for name in required}
            for row_number, row in enumerate(data_rows, 2):
                station = str(row[indexes['STATION']] or '').strip().upper()
                direction = str(row[indexes['DIRN']] or '').strip().upper()
                signal_type = str(row[indexes['TYPE']] or '').strip()
                if not station or not direction or not signal_type:
                    continue
                if direction.startswith('UP'):
                    event = 'J'
                    label = 'UP HOME'
                elif direction.startswith('DN'):
                    event = 'K'
                    label = 'DOWN HOME'
                else:
                    continue
                latitude = _number(row[indexes['LATITUDE']], f"{sheet.title} row {row_number} Latitude")
                longitude = _number(row[indexes['LONGITUDE']], f"{sheet.title} row {row_number} Longitude")
                if not -90 <= latitude <= 90 or not -180 <= longitude <= 180:
                    raise ValueError(f"{sheet.title} row {row_number}: coordinates are out of range.")
                key = f"{station}|{event}"
                mapping.setdefault(key, {'station': station, 'event': event, 'label': label, 'homes': []})['homes'].append({
                    'line': direction, 'type': signal_type, 'latitude': latitude, 'longitude': longitude,
                })
    finally:
        book.close()
    if not mapping:
        route_mapping = {}
        total = 0
        for sheet_name, rows in route_sheets:
            sheet_mapping, sheet_count = _parse_route_home_model_rows(rows, f"{filename} ({sheet_name})")
            for key, group in sheet_mapping.items():
                target = route_mapping.setdefault(key, {"station": group["station"], "event": group["event"],
                                                        "label": group["label"], "homes": []})
                seen = {(home.get("line", ""), home.get("type", ""), home.get("latitude"), home.get("longitude"))
                        for home in target["homes"]}
                for home in group["homes"]:
                    identity = (home.get("line", ""), home.get("type", ""), home.get("latitude"), home.get("longitude"))
                    if identity in seen:
                        continue
                    target["homes"].append(home)
                    seen.add(identity)
            total = sum(len(group["homes"]) for group in route_mapping.values())
        if not route_mapping:
            raise ValueError("No signal rows with UP/DN direction found.")
        return route_mapping, total
    return mapping, sum(len(value['homes']) for value in mapping.values())


def _haversine_m(latitude1, longitude1, latitude2, longitude2):
    radius = 6371008.8
    lat1, lat2 = math.radians(latitude1), math.radians(latitude2)
    dlat = lat2 - lat1
    dlon = math.radians(longitude2 - longitude1)
    value = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return round(2 * radius * math.asin(math.sqrt(min(1, value))), 1)


def _polygon_centroid(polygon):
    """Return the area-weighted centroid of (longitude, latitude) points."""
    area_twice = 0.0
    longitude_sum = 0.0
    latitude_sum = 0.0
    for first, second in zip(polygon, polygon[1:] + polygon[:1]):
        cross = first[0] * second[1] - second[0] * first[1]
        area_twice += cross
        longitude_sum += (first[0] + second[0]) * cross
        latitude_sum += (first[1] + second[1]) * cross
    if abs(area_twice) < 1e-12:
        return (sum(point[0] for point in polygon) / len(polygon),
                sum(point[1] for point in polygon) / len(polygon))
    return longitude_sum / (3 * area_twice), latitude_sum / (3 * area_twice)


def enrich_home_model(mapping, geofence_polygons):
    mapped = 0
    for value in mapping.values():
        polygon = geofence_polygons.get(value['station'], [])
        if polygon:
            station_lon, station_lat = _polygon_centroid(polygon)
            value['station_latitude'] = station_lat
            value['station_longitude'] = station_lon
            for home in value['homes']:
                home['distance_m'] = _haversine_m(station_lat, station_lon, home['latitude'], home['longitude'])
            mapped += 1
        else:
            for home in value['homes']:
                home['distance_m'] = None
    return mapped


def _merge_home_mappings(existing_mapping, incoming_mapping):
    """Union Home/I-Home signals without duplicating existing coordinates."""
    merged = json.loads(json.dumps(existing_mapping))
    for key, incoming in incoming_mapping.items():
        if key not in merged:
            merged[key] = incoming
            continue
        current = merged[key]
        seen = {
            (home.get('line', ''), home.get('type', ''), home.get('latitude'), home.get('longitude'))
            for home in current.get('homes', [])
        }
        for home in incoming.get('homes', []):
            identity = (home.get('line', ''), home.get('type', ''), home.get('latitude'), home.get('longitude'))
            if identity not in seen:
                current.setdefault('homes', []).append(home)
                seen.add(identity)
    return merged


def save_home_model(session: Session, filename: str, content: bytes, geofence_polygons=None, merge_existing: bool = True):
    if not filename.lower().endswith(('.xlsx', '.csv')):
        raise ValueError('Upload an .xlsx or .csv FSD file.')
    digest = sha256(content).hexdigest()
    existing = session.exec(select(RtisHomeModel).where(RtisHomeModel.digest == digest)).first()
    if existing:
        if geofence_polygons:
            mapping = json.loads(existing.mapping_json)
            mapped = enrich_home_model(mapping, geofence_polygons)
            existing.mapping_json = json.dumps(mapping)
            existing.mapped_station_count = mapped
            session.add(existing)
            session.commit()
            session.refresh(existing)
        return existing
    incoming_mapping, _ = parse_home_model(content, filename)
    previous_model = session.exec(select(RtisHomeModel).order_by(RtisHomeModel.id.desc())).first() if merge_existing else None
    if previous_model:
        mapping = _merge_home_mappings(json.loads(previous_model.mapping_json), incoming_mapping)
    else:
        mapping = incoming_mapping
    mapping, _ = _remove_known_bad_home_series(mapping)
    mapping, _ = _remove_nearby_duplicate_homes(mapping)
    if geofence_polygons:
        mapped = enrich_home_model(mapping, geofence_polygons)
    else:
        for value in mapping.values():
            for home in value.get('homes', []):
                home.setdefault('distance_m', None)
        mapped = sum(1 for value in mapping.values() if value.get('station_latitude') is not None)
    signal_count = sum(len(value.get('homes', [])) for value in mapping.values())
    model_name = filename.replace('\\', '/').split('/')[-1]
    if previous_model:
        model_name = f"Combined: {previous_model.filename} + {model_name}"
    model = RtisHomeModel(filename=model_name, digest=digest,
                           signal_count=signal_count, station_count=len({v['station'] for v in mapping.values()}),
                           mapped_station_count=mapped, mapping_json=json.dumps(mapping))
    try:
        session.add(model)
        session.commit()
        session.refresh(model)
    except IntegrityError:
        session.rollback()
        return session.exec(select(RtisHomeModel).where(RtisHomeModel.digest == digest)).one()
    return model


def selected_home_model(session: Session, model_id: int = 0):
    if model_id:
        model = session.get(RtisHomeModel, model_id)
    else:
        model = session.exec(select(RtisHomeModel).order_by(RtisHomeModel.id.desc())).first()
    if not model:
        return None, {}
    mapping, changed = _remove_known_bad_home_series(json.loads(model.mapping_json))
    mapping, duplicate_changed = _remove_nearby_duplicate_homes(mapping)
    changed = changed or duplicate_changed
    if changed:
        model.mapping_json = json.dumps(mapping)
        model.signal_count = sum(len(value.get('homes', [])) for value in mapping.values())
        model.station_count = len({value.get('station') for value in mapping.values() if value.get('station')})
        model.mapped_station_count = sum(1 for value in mapping.values() if value.get('station_latitude') is not None)
        session.add(model)
        session.commit()
        session.refresh(model)
    return model, mapping


def event_home_details(train_station, event_type, mapping):
    if event_type not in ('J', 'K'):
        return '', ''
    value = mapping.get(f"{str(train_station or '').strip().upper()}|{event_type}")
    if not value:
        return f"{'UP' if event_type == 'J' else 'DOWN'} HOME", ''
    distances = [home['distance_m'] for home in value['homes'] if home.get('distance_m') is not None]
    distance = min(distances) if distances else ''
    return value['label'], distance
