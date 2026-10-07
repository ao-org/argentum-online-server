# Copyright (C) 2026 Noland Studios LTD
# Licensed under the GNU Affero General Public License, version 3 or later.
"""Offline migration regressions using the supplied production row shape."""

import hashlib
from contextlib import closing
from pathlib import Path
import sqlite3
import struct
import tempfile
import unittest

from castle_migration import (
    CASTLE_COLUMNS, COORDINATE_COLUMNS, MIGRATION_FILE, MIGRATION_KEY, MigrationError, columns,
    convert_snapshot, map_bounds, migrate, rows,
)


ROOT = Path(__file__).resolve().parents[1]
BOUNDS = {number: (1, 100, 1, 100) for number in (27, 546, *range(758, 773))}


def fixture(connection):
    connection.executescript("CREATE TABLE account(id INTEGER PRIMARY KEY); CREATE TABLE user(id INTEGER PRIMARY KEY);")
    for name in ("20260724-02-regenerate castle table.sql", "20260724-05-rename castle column.sql",
                 "20260624-04-create castle coordinates.sql", "20260706-02-add tuple restrain on whitelist.sql"):
        if name.startswith("20260706"):
            connection.execute("CREATE TABLE castle_whitelist(id INTEGER)")
        connection.executescript((ROOT / "ScriptsDB" / name).read_text(encoding="cp1252"))
    connection.executemany("INSERT INTO account VALUES (?)", [(19241,), (9,), (4147,)])
    connection.executemany("INSERT INTO user VALUES (?)", [(15283,), (2897,), (1613,)])
    dates = ["2026-07-25 18:40:40", "2026-08-11 18:46:53.434", "2026-08-31 21:13:43"]
    # The first production name was clipped; use a fixture value, never invent its real name.
    names = ["Fixture castle name with spaces", None, "Thunder Hold"]
    for index in range(1, 21):
        account = [19241, 9, 4147][index - 1] if index <= 3 else None
        character = [15283, 2897, 1613][index - 1] if index <= 3 else None
        connection.execute("INSERT INTO castle VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                           (index, index + 20, account, character, index + 6361, index + 6382,
                            dates[index - 1] if index <= 3 else None, int(index <= 3),
                            names[index - 1] if index <= 3 else None))
    for index in range(1, 16):
        outside = (27, 49, 69) if index == 1 else (546, 50, 30) if index == 3 else (0, 0, 0)
        connection.execute("INSERT INTO castle_coordinates VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                           (index + 15, index, *outside, index + 757, 50, 72))
    connection.executemany("INSERT INTO castle_whitelist VALUES (?, ?, ?)",
                           [(101, "Owner Friend", 1), (102, "Pe\u00f1a", 2), (103, "MixedCase", 3)])
    connection.execute("UPDATE sqlite_sequence SET seq=50 WHERE name='castle'")
    connection.execute("UPDATE sqlite_sequence SET seq=80 WHERE name='castle_coordinates'")
    connection.commit()


class CastleMigrationTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.addCleanup(self.db.close)
        fixture(self.db)

    def test_production_shape_preserves_values_ids_and_missing_slots(self):
        original = rows(self.db, "castle", CASTLE_COLUMNS)
        whitelist = rows(self.db, "castle_whitelist", ("id", "character_name", "castle_id"))
        result = migrate(self.db, BOUNDS)
        self.assertEqual((result["castles"], result["coordinates"], result["entrance_bindings"]), (20, 15, 8))
        self.assertEqual(result["owned_unplaced_ids"], [2])
        self.assertEqual(result["unconfigured_ids"], [16, 17, 18, 19, 20])
        self.assertEqual(rows(self.db, "castle", CASTLE_COLUMNS), original)
        self.assertEqual(rows(self.db, "castle_whitelist", ("id", "character_name", "castle_id")), whitelist)
        self.assertEqual(self.db.execute("SELECT outside_map, outside_x, outside_y FROM castle_coordinates WHERE castle_id=2").fetchone(), (None, None, None))
        self.assertEqual(self.db.execute("SELECT foundation_date FROM castle WHERE id=2").fetchone()[0], "2026-08-11 18:46:53.434")
        self.assertEqual(self.db.execute("SELECT id, castle_id FROM castle_coordinates ORDER BY id").fetchall(), [(n + 15, n) for n in range(1, 16)])
        self.assertEqual(result["legacy_trigger_map"], {n + 20: n for n in range(1, 21)})
        self.assertEqual(self.db.execute("PRAGMA foreign_keys").fetchone()[0], 1)
        self.assertEqual(self.db.execute("PRAGMA foreign_key_check").fetchall(), [])

    def test_idempotence_preserves_normal_migration_history_and_data(self):
        migrate(self.db, BOUNDS)
        snapshot = list(self.db.iterdump())
        self.assertTrue(migrate(self.db, BOUNDS)["already_migrated"])
        self.assertEqual(list(self.db.iterdump()), snapshot)

    def test_non_contiguous_castle_ids_do_not_follow_coordinate_ids(self):
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.execute("UPDATE castle SET id=1001 WHERE id=1")
        self.db.execute("UPDATE castle SET id=87 WHERE id=3")
        self.db.commit()
        result = migrate(self.db, BOUNDS)
        self.assertEqual(result["legacy_trigger_map"][21], 1001)
        self.assertEqual(result["legacy_trigger_map"][23], 87)
        self.assertEqual({p["castle_id"] for p in result["placements"]}, {1001, 87})
        self.assertEqual(self.db.execute("SELECT castle_id FROM castle_coordinates WHERE id=16").fetchone()[0], 1001)

    def test_preserves_indexes_triggers_inbound_foreign_keys_and_sequences(self):
        self.db.executescript("""
            CREATE INDEX castle_name_index ON castle(name);
            CREATE TABLE audit(castle_id INTEGER);
            CREATE TRIGGER castle_name_audit AFTER UPDATE OF name ON castle BEGIN
                INSERT INTO audit VALUES(NEW.id);
            END;
            CREATE TABLE incoming(id INTEGER PRIMARY KEY, castle_id INTEGER REFERENCES castle(id));
            INSERT INTO incoming VALUES(1, 2);
        """)
        migrate(self.db, BOUNDS)
        self.assertIsNotNone(self.db.execute("SELECT sql FROM sqlite_master WHERE name='castle_name_index'").fetchone())
        self.db.execute("UPDATE castle SET name='Changed' WHERE id=2")
        self.assertEqual(self.db.execute("SELECT castle_id FROM audit").fetchall(), [(2,)])
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.execute("INSERT INTO incoming VALUES(2, 999)")
        castle = self.db.execute("INSERT INTO castle(spawner_obj_id, inside_key_obj_id) VALUES(7000, 7001)").lastrowid
        self.assertEqual(castle, 51)
        coordinate = self.db.execute("INSERT INTO castle_coordinates(castle_id, inside_map, inside_x, inside_y) VALUES(51, 758, 50, 72)").lastrowid
        self.assertEqual(coordinate, 81)

    def test_rejects_partial_zero_or_null_coordinates_and_rolls_back(self):
        for values in ((0, 49, 69), (27, None, 69), (None, 0, 0)):
            with self.subTest(values=values):
                self.db.execute("UPDATE castle_coordinates SET outside_map=?, outside_x=?, outside_y=? WHERE castle_id=1", values)
                self.db.commit()
                snapshot = list(self.db.iterdump())
                with self.assertRaises(MigrationError):
                    migrate(self.db, BOUNDS)
                self.assertEqual(list(self.db.iterdump()), snapshot)

    def test_rejects_out_of_bounds_full_footprints_and_duplicate_bindings(self):
        self.db.execute("UPDATE castle_coordinates SET outside_x=4 WHERE castle_id=1")
        self.db.commit()
        with self.assertRaisesRegex(MigrationError, "footprint"):
            migrate(self.db, BOUNDS)
        self.db.execute("UPDATE castle_coordinates SET outside_map=546,outside_x=50,outside_y=30 WHERE castle_id=1")
        self.db.commit()
        with self.assertRaisesRegex(MigrationError, "Overlapping"):
            migrate(self.db, BOUNDS)

    def test_rejects_duplicate_trigger_and_orphan_whitelist(self):
        self.db.execute('UPDATE castle SET "trigger"=21 WHERE id=2')
        self.db.commit()
        with self.assertRaisesRegex(MigrationError, "Duplicate legacy"):
            migrate(self.db, BOUNDS)
        self.db.execute('UPDATE castle SET "trigger"=22 WHERE id=2')
        self.db.execute("PRAGMA foreign_keys=OFF")
        self.db.commit()
        self.db.execute("PRAGMA foreign_keys=OFF")
        self.db.execute("INSERT INTO castle_whitelist VALUES(104, 'orphan', 999)")
        self.db.commit()
        with self.assertRaisesRegex(MigrationError, "foreign-key"):
            migrate(self.db, BOUNDS)

    def test_rejects_indexes_referencing_removed_column(self):
        self.db.execute('CREATE INDEX castle_trigger_index ON castle("trigger")')
        with self.assertRaisesRegex(MigrationError, "removed trigger"):
            migrate(self.db, BOUNDS)
        self.assertIn("trigger", columns(self.db, "castle"))

    def test_failure_during_sql_rolls_back_all_changes(self):
        self.db.execute("CREATE TABLE castle_legacy_trigger_map(do_not_replace INTEGER)")
        before = list(self.db.iterdump())
        with self.assertRaises(sqlite3.OperationalError):
            migrate(self.db, BOUNDS)
        self.assertEqual(list(self.db.iterdump()), before)
        self.assertEqual(self.db.execute("PRAGMA foreign_keys").fetchone()[0], 1)

    def test_target_constraint_rejects_partial_outside_coordinates(self):
        migrate(self.db, BOUNDS)
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.execute("UPDATE castle_coordinates SET outside_map=27 WHERE castle_id=2")

    def test_snapshot_mode_never_modifies_source_or_overwrites_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory) / "source.db", Path(directory) / "new.db"
            with closing(sqlite3.connect(source)) as file_db:
                self.db.backup(file_db)
            digest = hashlib.sha256(source.read_bytes()).digest()
            self.assertTrue(convert_snapshot(source, None, BOUNDS)["dry_run"])
            self.assertFalse(convert_snapshot(source, output, BOUNDS)["dry_run"])
            self.assertEqual(hashlib.sha256(source.read_bytes()).digest(), digest)
            with self.assertRaisesRegex(MigrationError, "already exists"):
                convert_snapshot(source, output, BOUNDS)
            with self.assertRaisesRegex(MigrationError, "different paths"):
                convert_snapshot(source, source, BOUNDS)
            with closing(sqlite3.connect(output)) as target:
                self.assertNotIn("trigger", columns(target, "castle"))

    def test_preserves_views_and_rejects_views_using_removed_trigger(self):
        self.db.execute("CREATE VIEW castle_names AS SELECT id, name FROM castle")
        expected = self.db.execute("SELECT * FROM castle_names ORDER BY id").fetchall()
        migrate(self.db, BOUNDS)
        self.assertEqual(self.db.execute("SELECT * FROM castle_names ORDER BY id").fetchall(), expected)
        other = sqlite3.connect(":memory:")
        self.addCleanup(other.close)
        fixture(other)
        other.execute('CREATE VIEW castle_triggers AS SELECT "trigger" AS legacy_id FROM castle')
        # Reject even quoted references: SQLite could otherwise turn a missing
        # double-quoted identifier into a string literal without raising an error.
        before = list(other.iterdump())
        with self.assertRaisesRegex(MigrationError, "removed trigger"):
            migrate(other, BOUNDS)
        self.assertEqual(list(other.iterdump()), before)

    def test_rejects_unknown_columns_and_preserves_other_migration_history(self):
        self.db.execute("ALTER TABLE castle ADD COLUMN custom_state TEXT")
        with self.assertRaisesRegex(MigrationError, "Unexpected castle columns"):
            migrate(self.db, BOUNDS)
        self.db.execute("ALTER TABLE castle DROP COLUMN custom_state")
        migrate(self.db, BOUNDS)
        self.db.execute("INSERT INTO migrations(date, description) VALUES ('20261008-01', 'unrelated later migration')")
        self.db.commit()
        self.assertTrue(migrate(self.db, BOUNDS)["already_migrated"])
        self.assertEqual(self.db.execute("SELECT date FROM migrations ORDER BY id").fetchall(),
                         [(MIGRATION_KEY,), ('20261008-01',)])

    def test_exact_scriptsdb_sql_with_runner_newline_removal(self):
        original = rows(self.db, "castle", CASTLE_COLUMNS)
        whitelist = rows(self.db, "castle_whitelist", ("id", "character_name", "castle_id"))
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.execute("CREATE TABLE entrance_audit(coordinate_id INTEGER REFERENCES castle_coordinates(id))")
        self.db.execute("INSERT INTO entrance_audit VALUES(16)")
        self.db.commit()
        script = MIGRATION_FILE.read_text(encoding="utf-8").replace("\r", "").replace("\n", "")
        self.db.executescript(script)
        self.assertEqual(rows(self.db, "castle", CASTLE_COLUMNS), original)
        self.assertEqual(rows(self.db, "castle_whitelist", ("id", "character_name", "castle_id")), whitelist)
        self.assertEqual(self.db.execute("SELECT * FROM entrance_audit").fetchall(), [(16,)])
        self.assertEqual(self.db.execute("SELECT outside_map,outside_x,outside_y FROM castle_coordinates WHERE castle_id=2").fetchone(), (None, None, None))
        self.assertEqual(self.db.execute("PRAGMA foreign_key_check").fetchall(), [])
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.execute("UPDATE castle_coordinates SET outside_map=27 WHERE castle_id=2")

    def test_scriptsdb_sql_failure_can_roll_back_entire_transaction(self):
        self.db.execute("UPDATE castle_coordinates SET outside_x=10 WHERE castle_id=2")
        self.db.commit()
        before = list(self.db.iterdump())
        with self.assertRaises(sqlite3.IntegrityError):
            self.db.executescript(MIGRATION_FILE.read_text(encoding="utf-8"))
        self.assertTrue(self.db.in_transaction)
        self.db.rollback()
        self.assertEqual(list(self.db.iterdump()), before)

    def test_reads_bounds_from_all_three_map_versions(self):
        with tempfile.TemporaryDirectory() as directory:
            for index, prefix, offset in ((1, b"", 44), (2, b"W5L2", 52), (3, b"CSM3\x02\x00\x05\x00", 68)):
                data = prefix + bytes(offset - len(prefix)) + struct.pack("<hhhh", 100, 1, 100, 1)
                (Path(directory) / f"Mapa{index}.csm").write_bytes(data)
            self.assertEqual(map_bounds(Path(directory)), {n: (1, 100, 1, 100) for n in range(1, 4)})

    def test_legacy_zero_bounds_normalize_but_modern_zero_bounds_fail(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "Mapa1.csm").write_bytes(b"W5L2" + bytes(56))
            self.assertEqual(map_bounds(root), {1: (1, 100, 1, 100)})
            (root / "Mapa2.csm").write_bytes(b"CSM3\x02\x00\x05\x00" + bytes(68))
            with self.assertRaisesRegex(MigrationError, "Invalid map bounds"):
                map_bounds(root)


if __name__ == "__main__":
    unittest.main()
