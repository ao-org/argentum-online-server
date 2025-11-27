#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Tool ReyarB - Activar Patron en Base de Datos (AO20)
# ----------------------------------------------------
# Actualiza la base SQLite del servidor.
#
# Ruta DB:
#   C:/Users/Administrador/Desktop/ServerAO20/argentum20-server/Database.db
#
# Tabla objetivo: account
# Columnas: is_active_patron, offline_patron_credits

import sqlite3
from pathlib import Path
import sys

DB_PATH = Path("C:\Users\alejo\Documents\GitHub\argentum-online-server")

IS_ACTIVE_VALUE = 6057395
OFFLINE_CREDITS = 150000


def error(msg: str, code: int = 1) -> None:
    print(f"ERROR: {msg}")
    sys.exit(code)


def resolve_account_id(cur, any_id: int):
    """
    Intenta resolver el ID ingresado como:
      1) account.id = any_id
      2) si no existe, user.id = any_id  -> devuelve user.account_id
    Retorna (account_id, origen_str) o (None, motivo)
    """
    # 1) account.id directo
    cur.execute("SELECT id FROM account WHERE id = ? LIMIT 1;", (any_id,))
    row = cur.fetchone()
    if row:
        return row[0], "account.id"

    # 2) user.id -> account_id
    cur.execute("SELECT account_id FROM user WHERE id = ? LIMIT 1;", (any_id,))
    row = cur.fetchone()
    if row and row[0] is not None:
        return int(row[0]), "user.id -> account_id"

    return None, "no encontrado como account.id ni como user.id"


def main() -> int:
    print("=== Activar Patron en Base de Datos (AO20) ===")
    print(f"Ruta DB: {DB_PATH}")

    if not DB_PATH.exists():
        error("No se encontro el archivo de base de datos en la ruta indicada.")

    raw = input("Ingrese ID (account.id o user.id): ").strip()
    if not raw.isdigit():
        error("ID invalido. Debe ser numerico.")

    any_id = int(raw)

    try:
        conn = sqlite3.connect(str(DB_PATH))
        cur = conn.cursor()

        account_id, origen = resolve_account_id(cur, any_id)
        if account_id is None:
            error("No se encontro ese ID ni en account.id ni en user.id.")

        print(f"ID resuelto -> account.id={account_id} (origen: {origen})")

        # Confirmar que existen las columnas esperadas
        cur.execute("PRAGMA table_info(account);")
        cols = {r[1].lower() for r in cur.fetchall()}
        for needed in ("is_active_patron", "offline_patron_credits"):
            if needed not in cols:
                error(f"Falta la columna '{needed}' en la tabla account.")

        # Actualizar
        cur.execute(
            """
            UPDATE account
               SET is_active_patron = ?,
                   offline_patron_credits = ?
             WHERE id = ?;
            """,
            (IS_ACTIVE_VALUE, OFFLINE_CREDITS, account_id),
        )
        conn.commit()

        if cur.rowcount > 0:
            print("OK: Actualizacion aplicada.")
            print(f"  account.id = {account_id}")
            print(f"  is_active_patron = {IS_ACTIVE_VALUE}")
            print(f"  offline_patron_credits = {OFFLINE_CREDITS}")
        else:
            print("Aviso: no se modifico ninguna fila (puede que ya tuviera esos valores).")

        return 0

    except sqlite3.Error as e:
        error(f"SQLite error: {e}")
    finally:
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado por el usuario.")
        sys.exit(130)
