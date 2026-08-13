"""Benutzer-Anmeldung für WAMAS Lift & Store Angebot.

users.json Format:
{
  "meta": {...},
  "users": [{ "username", "name", "email", "role", "salt", "hash", "active" }]
}
"""

from __future__ import annotations

import hashlib
import hmac
import json
import secrets
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from fastapi import APIRouter, HTTPException, Request, Response
from pydantic import BaseModel, Field

router = APIRouter(prefix="/api/auth", tags=["auth"])
SESSIONS: Dict[str, Dict[str, Any]] = {}

VALID_ROLES = {"viewer", "editor", "admin", "sales", "user"}


class UserLogin(BaseModel):
    username: str = Field(..., min_length=1)
    password: str = Field(..., min_length=1)


class UserCreate(BaseModel):
    username: str = Field(..., min_length=1)
    password: str = Field(..., min_length=1)
    name: Optional[str] = None
    displayName: Optional[str] = None
    email: Optional[str] = None
    role: str = "viewer"
    active: bool = True


class UserUpdate(BaseModel):
    name: Optional[str] = None
    displayName: Optional[str] = None
    email: Optional[str] = None
    password: Optional[str] = None
    role: Optional[str] = None
    active: Optional[bool] = None


def _users_path(base_dir: Path) -> Path:
    """Bevorzugt data/users.json (UI), Fallback Projektroot users.json."""
    data_path = base_dir / "data" / "users.json"
    root_path = base_dir / "users.json"
    if data_path.exists():
        return data_path
    if root_path.exists():
        return root_path
    # Neu anlegen im data/-Ordner (wie in der Benutzerverwaltung beschrieben)
    data_path.parent.mkdir(parents=True, exist_ok=True)
    return data_path


def _empty_store() -> Dict[str, Any]:
    return {
        "meta": {
            "note": "Rollen: viewer (lesen/export), editor (speichern), admin (löschen)."
        },
        "users": [],
    }


def _read_store(base_dir: Path) -> Dict[str, Any]:
    path = _users_path(base_dir)
    if not path.exists():
        store = _empty_store()
        salt, digest = hash_password("admin")
        store["users"].append(
            {
                "username": "admin",
                "name": "Administrator",
                "email": "",
                "role": "admin",
                "salt": salt,
                "hash": digest,
                "active": True,
            }
        )
        _write_store(base_dir, store)
        print("[Auth] users.json fehlte – Standard admin/admin angelegt.")
        return store
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return _empty_store()
    if isinstance(data, list):
        return {"meta": _empty_store()["meta"], "users": data}
    if not isinstance(data, dict):
        return _empty_store()
    data.setdefault("users", [])
    return data


def _write_store(base_dir: Path, store: Dict[str, Any]) -> None:
    _users_path(base_dir).write_text(
        json.dumps(store, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def hash_password(password: str, salt: Optional[str] = None) -> Tuple[str, str]:
    if not salt:
        salt = secrets.token_hex(16)
    return salt, _digest_primary(password, salt)


def _digest_primary(password: str, salt: str) -> str:
    try:
        salt_bytes = bytes.fromhex(salt)
    except ValueError:
        salt_bytes = salt.encode("utf-8")
    return hashlib.sha256(salt_bytes + password.encode("utf-8")).hexdigest()


def _candidate_digests(password: str, salt: str) -> List[str]:
    out: List[str] = []
    pw = password.encode("utf-8")
    try:
        salt_bytes = bytes.fromhex(salt)
    except ValueError:
        salt_bytes = None
    salt_str = salt.encode("utf-8")
    out.append(_digest_primary(password, salt))
    out.append(hashlib.sha256(salt_str + pw).hexdigest())
    out.append(hashlib.sha256(pw + salt_str).hexdigest())
    out.append(hashlib.sha256(f"{salt}:{password}".encode("utf-8")).hexdigest())
    out.append(hashlib.sha256(f"{password}:{salt}".encode("utf-8")).hexdigest())
    out.append(hmac.new(salt_str, pw, hashlib.sha256).hexdigest())
    if salt_bytes is not None:
        out.append(hashlib.sha256(salt_bytes + pw).hexdigest())
        out.append(hashlib.sha256(pw + salt_bytes).hexdigest())
        out.append(hmac.new(salt_bytes, pw, hashlib.sha256).hexdigest())
        for rounds in (1000, 10000, 100000, 120000):
            out.append(hashlib.pbkdf2_hmac("sha256", pw, salt_bytes, rounds).hex())
            out.append(hashlib.pbkdf2_hmac("sha256", pw, salt_str, rounds).hex())
    return out


def verify_password(password: str, salt: str, expected_hash: str) -> bool:
    if not salt or not expected_hash:
        return False
    expected = expected_hash.lower().strip()
    for cand in _candidate_digests(password, salt):
        if hmac.compare_digest(cand.lower(), expected):
            return True
    return False


def _norm_role(role: Any) -> str:
    r = str(role or "viewer").strip().lower()
    if r in {"administrator", "adm"}:
        return "admin"
    if r in {"benutzer", "user"}:
        return "user"
    if r in {"verkauf", "sale", "vertrieb"}:
        return "sales"
    if r in {"lesen", "read", "readonly"}:
        return "viewer"
    return r


def _is_admin(session: Optional[Dict[str, Any]]) -> bool:
    if not session:
        return False
    return _norm_role(session.get("role")) == "admin"


def _public_user(u: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "username": u.get("username"),
        "name": u.get("name") or u.get("displayName") or u.get("username"),
        "displayName": u.get("name") or u.get("displayName") or u.get("username"),
        "email": u.get("email") or "",
        "role": _norm_role(u.get("role")),
        "active": bool(u.get("active", True)),
    }


def _find_user(store: Dict[str, Any], username: str) -> Optional[Dict[str, Any]]:
    uname = (username or "").strip().lower()
    for u in store.get("users") or []:
        if str(u.get("username") or "").strip().lower() == uname:
            return u
    return None


def _authenticate(base_dir: Path, username: str, password: str) -> Optional[Dict[str, Any]]:
    store = _read_store(base_dir)
    user = _find_user(store, username)
    if not user or user.get("active") is False:
        return None
    if user.get("salt") and user.get("hash"):
        if verify_password(password, str(user["salt"]), str(user["hash"])):
            return user
        return None
    legacy = user.get("password")
    if legacy:
        simple = hashlib.sha256(password.encode("utf-8")).hexdigest()
        if hmac.compare_digest(str(legacy).lower(), simple.lower()):
            return user
    return None


def _create_session(user: Dict[str, Any]) -> str:
    token = secrets.token_hex(32)
    SESSIONS[token] = {
        "username": user.get("username"),
        "name": user.get("name") or user.get("displayName") or user.get("username"),
        "displayName": user.get("name") or user.get("displayName") or user.get("username"),
        "email": user.get("email") or "",
        "role": _norm_role(user.get("role")),
        "created": time.time(),
    }
    return token


def _token_from_request(request: Request) -> str:
    """Token aus allen üblichen Stellen lesen (Header/Cookie/Query)."""
    auth = (request.headers.get("Authorization") or "").strip()
    if auth.lower().startswith("bearer "):
        return auth[7:].strip()
    if auth.lower().startswith("token "):
        return auth[6:].strip()
    if auth and " " not in auth:
        return auth

    for key in (
        "X-Auth-Token",
        "x-auth-token",
        "X-Session-Token",
        "Auth-Token",
        "Token",
    ):
        val = (request.headers.get(key) or "").strip()
        if val:
            return val

    for key in ("auth_token", "token", "session", "wamas_auth_token"):
        val = (request.cookies.get(key) or "").strip()
        if val:
            return val

    q = (request.query_params.get("token") or request.query_params.get("auth") or "").strip()
    if q:
        return q
    return ""


def _session(request: Request) -> Optional[Dict[str, Any]]:
    tok = _token_from_request(request)
    if not tok:
        return None
    return SESSIONS.get(tok)


def _set_auth_cookie(response: Response, token: str) -> None:
    response.set_cookie(
        key="auth_token",
        value=token,
        httponly=False,  # Frontend darf Token auch aus Cookie lesen, falls nötig
        samesite="lax",
        path="/",
        max_age=60 * 60 * 12,
    )


def _clear_auth_cookie(response: Response) -> None:
    response.delete_cookie("auth_token", path="/")


def _require_admin(request: Request) -> Dict[str, Any]:
    session = _session(request)
    if not _is_admin(session):
        # Hilfreiche Diagnose im Server-Log
        tok = _token_from_request(request)
        print(
            "[Auth] 403 Benutzerverwaltung – "
            f"token_present={bool(tok)} session={bool(session)} "
            f"role={(session or {}).get('role')!r}"
        )
        raise HTTPException(status_code=403, detail="Nur Administratoren.")
    return session or {}


def create_auth_router(base_dir: Path) -> APIRouter:
    _read_store(base_dir)

    @router.post("/login")
    def auth_login(body: UserLogin, response: Response):
        user = _authenticate(base_dir, body.username, body.password)
        if not user:
            raise HTTPException(status_code=401, detail="Benutzername oder Passwort falsch.")
        token = _create_session(user)
        pub = _public_user(user)
        _set_auth_cookie(response, token)
        return {
            "ok": True,
            "token": token,
            "username": pub["username"],
            "name": pub["name"],
            "displayName": pub["name"],
            "email": pub["email"],
            "role": pub["role"],
        }

    @router.post("/logout")
    def auth_logout(request: Request, response: Response):
        tok = _token_from_request(request)
        SESSIONS.pop(tok, None)
        _clear_auth_cookie(response)
        return {"ok": True}

    @router.get("/session")
    def auth_session(request: Request):
        session = _session(request)
        if not session:
            return {"authenticated": False}
        return {
            "authenticated": True,
            "username": session.get("username"),
            "name": session.get("name"),
            "displayName": session.get("displayName"),
            "email": session.get("email"),
            "role": session.get("role"),
        }

    @router.get("/me")
    def auth_me(request: Request):
        return auth_session(request)

    @router.get("/users")
    def list_users(request: Request):
        _require_admin(request)
        store = _read_store(base_dir)
        return {"users": [_public_user(u) for u in store.get("users") or []]}

    @router.post("/users")
    def create_user(body: UserCreate, request: Request):
        _require_admin(request)
        store = _read_store(base_dir)
        if _find_user(store, body.username):
            raise HTTPException(status_code=409, detail=f"Benutzer '{body.username}' existiert bereits.")
        role = _norm_role(body.role)
        if role not in VALID_ROLES:
            raise HTTPException(status_code=400, detail=f"Ungültige Rolle. Erlaubt: {sorted(VALID_ROLES)}")
        if len(body.password or "") < 6:
            raise HTTPException(status_code=400, detail="Passwort mindestens 6 Zeichen.")
        salt, digest = hash_password(body.password)
        display = (body.name or body.displayName or body.username).strip()
        store["users"].append(
            {
                "username": body.username.strip(),
                "name": display,
                "email": (body.email or "").strip(),
                "role": role,
                "salt": salt,
                "hash": digest,
                "active": bool(body.active),
            }
        )
        _write_store(base_dir, store)
        return {"ok": True, "message": f"Benutzer '{body.username}' erstellt.", "user": _public_user(store["users"][-1])}

    def _update_user_impl(username: str, body: UserUpdate, request: Request):
        _require_admin(request)
        store = _read_store(base_dir)
        user = _find_user(store, username)
        if not user:
            raise HTTPException(status_code=404, detail="Benutzer nicht gefunden.")
        if body.name is not None or body.displayName is not None:
            user["name"] = (body.name or body.displayName or "").strip()
        if body.email is not None:
            user["email"] = body.email.strip()
        if body.role is not None:
            role = _norm_role(body.role)
            if role not in VALID_ROLES:
                raise HTTPException(status_code=400, detail=f"Ungültige Rolle. Erlaubt: {sorted(VALID_ROLES)}")
            user["role"] = role
        if body.active is not None:
            user["active"] = bool(body.active)
        if body.password is not None and body.password.strip():
            if len(body.password.strip()) < 6:
                raise HTTPException(status_code=400, detail="Passwort mindestens 6 Zeichen.")
            salt, digest = hash_password(body.password.strip())
            user["salt"] = salt
            user["hash"] = digest
            user.pop("password", None)
        _write_store(base_dir, store)
        return {"ok": True, "user": _public_user(user)}

    @router.put("/users/{username}")
    def update_user(username: str, body: UserUpdate, request: Request):
        return _update_user_impl(username, body, request)

    # Viele Frontends senden PATCH statt PUT → sonst HTTP 405 Method Not Allowed
    @router.patch("/users/{username}")
    def patch_user(username: str, body: UserUpdate, request: Request):
        return _update_user_impl(username, body, request)

    class PasswordChange(BaseModel):
        password: str = Field(..., min_length=6)
        passwordConfirm: Optional[str] = None

    def _set_password_impl(username: str, body: PasswordChange, request: Request):
        _require_admin(request)
        if body.passwordConfirm is not None and body.passwordConfirm != body.password:
            raise HTTPException(status_code=400, detail="Passwörter stimmen nicht überein.")
        return _update_user_impl(
            username,
            UserUpdate(password=body.password),
            request,
        )

    @router.put("/users/{username}/password")
    def put_password(username: str, body: PasswordChange, request: Request):
        return _set_password_impl(username, body, request)

    @router.patch("/users/{username}/password")
    def patch_password(username: str, body: PasswordChange, request: Request):
        return _set_password_impl(username, body, request)

    @router.post("/users/{username}/password")
    def post_password(username: str, body: PasswordChange, request: Request):
        return _set_password_impl(username, body, request)

    @router.delete("/users/{username}")
    def delete_user(username: str, request: Request):
        _require_admin(request)
        if username.strip().lower() == "admin":
            raise HTTPException(status_code=400, detail="Standard-Admin kann nicht gelöscht werden.")
        store = _read_store(base_dir)
        before = len(store.get("users") or [])
        store["users"] = [
            u
            for u in (store.get("users") or [])
            if str(u.get("username") or "").strip().lower() != username.strip().lower()
        ]
        if len(store["users"]) == before:
            raise HTTPException(status_code=404, detail="Benutzer nicht gefunden.")
        _write_store(base_dir, store)
        return {"ok": True, "message": f"Benutzer '{username}' gelöscht."}

    return router
