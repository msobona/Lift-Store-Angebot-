#!/usr/bin/env python3
"""Stellt bekannte Benutzer wieder her (ohne alte Passwort-Hashes).

Setzt für msobona / caytac / admin neue Passwörter, die du angibst.
Bestehende users.json wird aktualisiert (nicht blind gelöscht).

Beispiel:
  python scripts/restore_known_users.py
  (fragt interaktiv nach Passwörtern)

oder:
  python scripts/restore_known_users.py adminPass msobonaPass caytacPass
"""

from __future__ import annotations

import getpass
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from auth import hash_password  # noqa: E402

USERS_SPEC = [
    {"username": "admin", "name": "Administrator", "email": "", "role": "admin"},
    {
        "username": "msobona",
        "name": "Marek Sobona",
        "email": "marek.sobona@ssi-schaefer.com",
        "role": "admin",  # gewünschte Rolle laut UI; vorher viewer
    },
    {"username": "caytac", "name": "Cenk", "email": "", "role": "sales"},
]


def main() -> int:
    path = ROOT / "users.json"
    if len(sys.argv) >= 4:
        passwords = {
            "admin": sys.argv[1],
            "msobona": sys.argv[2],
            "caytac": sys.argv[3],
        }
    else:
        print("Passwörter setzen (min. 6 Zeichen). Bestehende Hashes werden ersetzt.")
        passwords = {}
        for spec in USERS_SPEC:
            pw = getpass.getpass(f"Passwort für {spec['username']}: ").strip()
            if len(pw) < 6:
                print("Zu kurz.")
                return 1
            passwords[spec["username"]] = pw

    store = {
        "meta": {
            "note": "Rollen: viewer (lesen/export), editor (speichern), admin (löschen)."
        },
        "users": [],
    }
    for spec in USERS_SPEC:
        salt, digest = hash_password(passwords[spec["username"]])
        store["users"].append(
            {
                **spec,
                "salt": salt,
                "hash": digest,
                "active": True,
            }
        )
    path.write_text(json.dumps(store, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("OK – users.json geschrieben:", path)
    print("Benutzer:", ", ".join(u["username"] for u in store["users"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
