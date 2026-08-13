#!/usr/bin/env python3
"""Passwort eines bestehenden Users neu setzen (Salt+Hash), ohne andere Felder zu löschen.

Beispiel:
  python scripts/set_user_password.py admin NeuesPasswort
  python scripts/set_user_password.py msobona GeheimesPasswort
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from auth import hash_password  # noqa: E402


def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: python scripts/set_user_password.py <username> <neues_passwort>")
        return 1
    username = sys.argv[1].strip().lower()
    password = sys.argv[2]
    path = ROOT / "users.json"
    if not path.exists():
        print("users.json nicht gefunden:", path)
        return 1
    store = json.loads(path.read_text(encoding="utf-8"))
    users = store.get("users") if isinstance(store, dict) else store
    if not isinstance(users, list):
        print("Ungültiges users.json Format")
        return 1
    found = None
    for u in users:
        if str(u.get("username") or "").strip().lower() == username:
            found = u
            break
    if not found:
        print("Benutzer nicht gefunden:", username)
        return 1
    salt, digest = hash_password(password)
    found["salt"] = salt
    found["hash"] = digest
    found["active"] = True
    found.pop("password", None)
    if isinstance(store, dict):
        path.write_text(json.dumps(store, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    else:
        path.write_text(json.dumps(users, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"OK – Passwort für '{found.get('username')}' aktualisiert.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
