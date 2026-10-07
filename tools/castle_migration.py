# Copyright (C) 2026 Noland Studios LTD
# Licensed under the GNU Affero General Public License, version 3 or later.
"""Migrate a read-only castle database snapshot to a separately validated output."""

from __future__ import annotations

import argparse
from contextlib import closing
import json
import os
from pathlib import Path
import re
import sqlite3
import struct
import tempfile


MIGRATION_KEY = "20261007-01"
MIGRATION_FILE = Path(__file__).resolve().parents[1] / "ScriptsDB" / "20261007-01-migrate castle identities.sql"
CASTLE_COLUMNS = (
    "id", "owner_account_id", "owner_character_id", "spawner_obj_id",
    "inside_key_obj_id", "foundation_date", "is_active", "name",
)
COORDINATE_COLUMNS = (
    "id", "castle_id", "outside_map", "outside_x", "outside_y",
    "inside_map", "inside_x", "inside_y",
)


class MigrationError(RuntimeError):
    pass


def quote(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def columns(connection: sqlite3.Connection, table: str) -> tuple[str, ...]:
    return tuple(row[1] for row in connection.execute(f"PRAGMA table_info({quote(table)})"))


def rows(connection: sqlite3.Connection, table: str, names: tuple[str, ...]) -> list[tuple]:
    fields = ", ".join(map(quote, names))
    return connection.execute(f"SELECT {fields} FROM {quote(table)} ORDER BY id").fetchall()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise MigrationError(message)


def integer(value: object, label: str, maximum: int = 2147483647) -> int:
    require(type(value) is int and 0 < value <= maximum, f"Invalid positive integer {label}: {value!r}")
    return value


def map_bounds(directory: Path) -> dict[int, tuple[int, int, int, int]]:
    """Read only the bounds shared by legacy four/five-layer maps and CSM3."""
    result = {}
    for path in directory.glob("*.csm"):
        match = re.fullmatch(r"(?:mapa)?(\d+)", path.stem, re.IGNORECASE)
        require(match is not None, f"Cannot determine map ID from {path.name}")
        map_id = int(match[1])
        with path.open("rb") as stream:
            header = stream.read(76)
        offset = 68 if header[:4] == b"CSM3" else 52 if header[:4] == b"W5L2" else 44
        require(len(header) >= offset + 8, f"Truncated map header: {path}")
        if header[:4] == b"CSM3":
            require(struct.unpack_from("<HH", header, 4) == (2, 5), f"Unsupported CSM3 version: {path}")
        xmax, xmin, ymax, ymin = struct.unpack_from("<hhhh", header, offset)
        if header[:4] != b"CSM3" and (xmax, xmin, ymax, ymin) == (0, 0, 0, 0):
            xmax, xmin, ymax, ymin = 100, 1, 100, 1
        require(1 <= xmin <= xmax <= 255 and 1 <= ymin <= ymax <= 255,
                f"Invalid map bounds: {path}")
        require(map_id not in result, f"Duplicate map ID: {map_id}")
        result[map_id] = xmin, xmax, ymin, ymax
    require(bool(result), f"No map files found in {directory}")
    return result


def validate(connection: sqlite3.Connection, bounds: dict[int, tuple[int, int, int, int]],
             legacy: bool) -> dict:
    expected = set(CASTLE_COLUMNS) | ({"trigger"} if legacy else set())
    require(set(columns(connection, "castle")) == expected,
            "Unexpected castle columns; review the actual schema before migration")
    require(set(columns(connection, "castle_coordinates")) == set(COORDINATE_COLUMNS),
            "Unexpected castle_coordinates columns; review the actual schema before migration")
    require(columns(connection, "castle_whitelist") == ("id", "character_name", "castle_id"),
            "Missing or unexpected castle_whitelist schema")
    require(not connection.execute("PRAGMA foreign_key_check").fetchall(),
            "Database has foreign-key violations; correct them before migration")
    castle_rows = rows(connection, "castle", CASTLE_COLUMNS)
    castle_ids = {row[0] for row in castle_rows}
    for record in castle_rows:
        integer(record[0], "castle.id")
        for index in (1, 2):
            if record[index] is not None:
                integer(record[index], CASTLE_COLUMNS[index])
        for index in (3, 4):
            integer(record[index], CASTLE_COLUMNS[index], 32767)
        require(record[6] in (0, 1), f"Invalid castle activity: {record[0]}")
    for index in (1, 2, 3, 4):
        values = [row[index] for row in castle_rows if row[index] is not None]
        require(len(values) == len(set(values)), f"Duplicate {CASTLE_COLUMNS[index]}")
    mapping = connection.execute('SELECT "trigger", id FROM castle ORDER BY "trigger"').fetchall() if legacy else (
        connection.execute("SELECT legacy_trigger, castle_id FROM castle_legacy_trigger_map ORDER BY legacy_trigger").fetchall())
    require(len(mapping) == len({row[0] for row in mapping}), "Duplicate legacy castle trigger")
    for trigger, castle_id in mapping:
        require(type(trigger) is int and 21 <= trigger <= 40, f"Unknown legacy castle trigger: {trigger}")
        require(castle_id in castle_ids, f"Orphan legacy castle mapping: {castle_id}")
    placements = []
    configured = set()
    entrance_tiles = set()
    for row in rows(connection, "castle_coordinates", COORDINATE_COLUMNS):
        coordinate_id, castle_id, omap, ox, oy, imap, ix, iy = row
        integer(coordinate_id, "castle_coordinates.id")
        require(castle_id in castle_ids, f"Orphan coordinate row: {coordinate_id}")
        require(castle_id not in configured, f"Duplicate castle coordinates: {castle_id}")
        configured.add(castle_id)
        for map_id, x, y, inside in ((imap, ix, iy, True), (omap, ox, oy, False)):
            if not inside and ((map_id, x, y) == (None, None, None) or (legacy and (map_id, x, y) == (0, 0, 0))):
                continue
            integer(map_id, "map", 32767)
            integer(x, "x", 255)
            integer(y, "y", 255)
            require(map_id in bounds, f"Map {map_id} has no supplied bounds (castle {castle_id})")
            xmin, xmax, ymin, ymax = bounds[map_id]
            require(xmin <= x <= xmax and ymin <= y <= ymax, f"Coordinates outside map {map_id} (castle {castle_id})")
            if inside:
                continue
            require(xmin <= x - 8 and x + 6 <= xmax and ymin <= y - 8 and y + 2 <= ymax,
                    f"Castle {castle_id} footprint crosses map bounds")
            placements.append({"castle_id": castle_id, "map": map_id, "x": x, "y": y})
            for tx, ty in ((x - 1, y), (x - 2, y), (x - 1, y + 1), (x - 2, y + 1)):
                key = map_id, tx, ty
                require(key not in entrance_tiles, f"Overlapping castle entrance: {key}")
                entrance_tiles.add(key)
    placed = {row["castle_id"] for row in placements}
    whitelist = rows(connection, "castle_whitelist", ("id", "character_name", "castle_id"))
    require(all(row[2] in castle_ids for row in whitelist), "Orphan whitelist entry")
    return {
        "migration": MIGRATION_KEY, "castles": len(castle_rows), "coordinates": len(configured),
        "whitelist_rows": len(whitelist), "legacy_trigger_map": dict(mapping),
        "placements": placements, "entrance_bindings": len(entrance_tiles),
        "owned_unplaced_ids": [row[0] for row in castle_rows if row[1] is not None and row[0] not in placed],
        "unconfigured_ids": sorted(castle_ids - configured),
    }


def migrate(connection: sqlite3.Connection, bounds: dict[int, tuple[int, int, int, int]]) -> dict:
    require(not connection.in_transaction, "Migration requires a connection outside a transaction")
    require(sqlite3.sqlite_version_info >= (3, 35, 0), "Castle SQL migration requires SQLite 3.35 or newer")
    connection.execute("PRAGMA foreign_keys=ON")
    applied = bool(columns(connection, "migrations")) and connection.execute(
        "SELECT 1 FROM migrations WHERE date = ?", (MIGRATION_KEY,)).fetchone()
    if applied:
        report = validate(connection, bounds, legacy=False)
        report["already_migrated"] = True
        return report
    before = validate(connection, bounds, legacy=True)
    # Check the documented production constraints before rehearsing the SQL.
    expected_unique = {
        "castle": {("owner_account_id",), ("owner_character_id",), ("spawner_obj_id",), ("inside_key_obj_id",)},
        "castle_coordinates": {("castle_id",)},
    }
    expected_foreign = {
        "castle": {("account", "owner_account_id", "id", "CASCADE", "CASCADE"),
                   ("user", "owner_character_id", "id", "CASCADE", "CASCADE")},
        "castle_coordinates": {("castle", "castle_id", "id", "CASCADE", "CASCADE")},
    }
    for table in expected_unique:
        sql = connection.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()[0]
        require(not re.search(r"\b(CHECK|COLLATE|GENERATED|WITHOUT|STRICT)\b", sql, re.IGNORECASE),
                f"Unrecognized constraints on {table}; review before migration")
        uniques = {tuple(part[2] for part in connection.execute(f"PRAGMA index_info({quote(index[1])})"))
                   for index in connection.execute(f"PRAGMA index_list({quote(table)})")
                   if index[2] and index[3] == "u"}
        require(uniques == expected_unique[table], f"Unexpected unique constraints on {table}")
        foreign = {tuple(row[2:7]) for row in connection.execute(f"PRAGMA foreign_key_list({quote(table)})")}
        require(foreign == expected_foreign[table], f"Unexpected foreign keys on {table}")
    castle_rows = rows(connection, "castle", CASTLE_COLUMNS)
    coordinate_rows = [row[:2] + ((None, None, None) if row[2:5] == (0, 0, 0) else row[2:5]) + row[5:]
                       for row in rows(connection, "castle_coordinates", COORDINATE_COLUMNS)]
    whitelist_rows = rows(connection, "castle_whitelist", ("id", "character_name", "castle_id"))
    sequences = dict(connection.execute("SELECT name, seq FROM sqlite_sequence WHERE name IN ('castle', 'castle_coordinates')"))
    schema_objects = connection.execute("SELECT type, name, sql FROM sqlite_master WHERE tbl_name IN ('castle', 'castle_coordinates') AND type IN ('index', 'trigger') AND sql IS NOT NULL").fetchall()
    views = connection.execute("SELECT name, sql FROM sqlite_master WHERE type='view'").fetchall()
    for name, sql in views:
        require(not (re.search(r"\bcastle\b", sql, re.IGNORECASE) and re.search(r"\btrigger\b", sql, re.IGNORECASE)),
                f"View {name} may reference removed trigger column; migrate it explicitly")
    for kind, name, sql in schema_objects:
        occurrences = len(re.findall(r"\btrigger\b", sql, re.IGNORECASE))
        require(occurrences <= (1 if kind == "trigger" else 0),
                f"Schema object {name} may reference removed trigger column; migrate it explicitly")
    try:
        # Execute the dated SQL itself. Delay its COMMIT until the rehearsal's
        # preservation checks pass, then record the normal runner history row.
        pending = ""
        found_commit = False
        for line in MIGRATION_FILE.read_text(encoding="utf-8").splitlines(keepends=True):
            pending += line
            if sqlite3.complete_statement(pending):
                if pending.strip().upper() == "COMMIT;":
                    found_commit = True
                else:
                    require(not found_commit, "Unexpected SQL after migration COMMIT")
                    connection.execute(pending)
                pending = ""
        require(found_commit and not pending.strip() and connection.in_transaction,
                "Migration SQL must finish its explicit transaction with COMMIT")
        require(rows(connection, "castle", CASTLE_COLUMNS) == castle_rows, "Castle preservation check failed")
        require(rows(connection, "castle_coordinates", COORDINATE_COLUMNS) == coordinate_rows, "Coordinate preservation check failed")
        require(rows(connection, "castle_whitelist", ("id", "character_name", "castle_id")) == whitelist_rows, "Whitelist preservation check failed")
        require(dict(connection.execute("SELECT name, seq FROM sqlite_sequence WHERE name IN ('castle', 'castle_coordinates')")) == sequences,
                "Autoincrement sequence preservation check failed")
        for (name,) in connection.execute("SELECT name FROM sqlite_master WHERE type='view'"):
            connection.execute(f"SELECT * FROM {quote(name)} LIMIT 0")
        after = validate(connection, bounds, legacy=False)
        require(after == before, "Migration changed castle placement or reference inventory")
        require(connection.execute("PRAGMA integrity_check").fetchall() == [("ok",)], "Database integrity check failed")
        connection.execute('CREATE TABLE IF NOT EXISTS migrations (id INTEGER NOT NULL PRIMARY KEY, date VARCHAR(11) NOT NULL, description VARCHAR(50) NULL)')
        connection.execute("INSERT INTO migrations(date, description) VALUES (?, ?)",
                           (MIGRATION_KEY, "migrate castle identities"))
        connection.commit()
    except BaseException:
        connection.rollback()
        raise
    after["already_migrated"] = False
    return after


def convert_snapshot(source: Path, output: Path | None, bounds: dict) -> dict:
    source = source.resolve(strict=True)
    require(source.is_file(), "Source must be a SQLite database file")
    if output is not None:
        output = output.absolute()
        require(output.resolve() != source, "Source and output must be different paths")
        require(not output.exists(), f"Output already exists: {output}")
        require(output.parent.is_dir(), "Output parent directory must already exist")
    with tempfile.TemporaryDirectory(prefix="castle-migration-", dir=output.parent if output else None) as scratch:
        candidate = Path(scratch) / "migrated.db"
        with closing(sqlite3.connect(source.as_uri() + "?mode=ro", uri=True)) as original:
            with closing(sqlite3.connect(candidate)) as target:
                original.backup(target)
                report = migrate(target, bounds)
        if output is not None:
            # Atomic publication without replacing an existing file, even in a race.
            os.link(candidate, output)
    report["dry_run"] = output is None
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True, help="Consistent offline SQLite snapshot; never modified")
    parser.add_argument("--maps", type=Path, required=True, help="Map directory used to validate interiors and full outside footprints")
    destination = parser.add_mutually_exclusive_group(required=True)
    destination.add_argument("--dry-run", action="store_true")
    destination.add_argument("--output", type=Path, help="New database path; existing files are never overwritten")
    args = parser.parse_args()
    try:
        report = convert_snapshot(args.source, args.output, map_bounds(args.maps))
    except (MigrationError, sqlite3.Error, OSError) as error:
        parser.exit(1, f"Castle migration failed: {error}\n")
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
