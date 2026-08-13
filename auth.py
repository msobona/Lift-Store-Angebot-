"""Benutzer-Anmeldung für WAMAS Lift & Store Angebot."""

from __future__ import annotations

import hashlib
import json
import secrets
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel, Field

router = APIRouter(prefix="/api/auth", tags=["auth"])

SESSIONS: Dict[str, Dict[str, Any]] = {}


class UserLogin(BaseModel):
    username: str = Field(..., min_length=1)
    password: str = Field(..., min_length=1)


class UserCreate(BaseModel):
    username: str = Field(..., min_length=1)
    password: str = Field(..., min_length=1)
    displayName: Optional[str] = None
    role: str = "user"


class UserUpdate(BaseModel):
    displayName: Optional[str] = None
    password: Optional[str] = None
    role: Optional[str] = None


def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _users_file(base_dir: Path) -> Path:
    return base_dir / "users.json"


def init_users(base_dir: Path) -> None:
    """Legt Standard-Admin an, falls users.json fehlt."""
    path = _users_file(base_dir)
    if path.exists():
        return
    default_users = [
        {
            "username": "admin",
            "password": _hash_password("admin"),
            "displayName": "Administrator",
            "role": "admin",
        }
    ]
    path.write_text(json.dumps(default_users, ensure_ascii=False, indent=2), encoding="utf-8")
    print("[Auth] Standard-Benutzer 'admin' erstellt (Passwort: admin)")


def _read_users(base_dir: Path) -> List[Dict[str, Any]]:
    init_users(base_dir)
    try:
        return json.loads(_users_file(base_dir).read_text(encoding="utf-8"))
    except Exception:
        return []


def _write_users(base_dir: Path, users: List[Dict[str, Any]]) -> None:
    _users_file(base_dir).write_text(
        json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _find_user(base_dir: Path, username: str) -> Optional[Dict[str, Any]]:
    for u in _read_users(base_dir):
        if u.get("username") == username:
            return u
    return None


def _authenticate(base_dir: Path, username: str, password: str) -> Optional[Dict[str, Any]]:
    user = _find_user(base_dir, username)
    if user and user.get("password") == _hash_password(password):
        return user
    return None


def _create_session(user: Dict[str, Any]) -> str:
    token = secrets.token_hex(32)
    SESSIONS[token] = {
        "username": user["username"],
        "displayName": user.get("displayName") or user["username"],
        "role": user.get("role") or "user",
        "created": time.time(),
    }
    return token


def _get_session(token: str) -> Optional[Dict[str, Any]]:
    if not token:
        return None
    return SESSIONS.get(token)


def _destroy_session(token: str) -> None:
    SESSIONS.pop(token, None)


def _token_from_request(request: Request) -> str:
    return (
        request.headers.get("X-Auth-Token")
        or request.headers.get("Authorization", "").removeprefix("Bearer ").strip()
        or request.cookies.get("auth_token")
        or ""
    )


def create_auth_router(base_dir: Path) -> APIRouter:
    init_users(base_dir)

    @router.post("/login")
    def auth_login(body: UserLogin):
        user = _authenticate(base_dir, body.username.strip(), body.password)
        if not user:
            raise HTTPException(status_code=401, detail="Benutzername oder Passwort falsch.")
        token = _create_session(user)
        return {
            "ok": True,
            "token": token,
            "username": user["username"],
            "displayName": user.get("displayName") or user["username"],
            "role": user.get("role") or "user",
        }

    @router.post("/logout")
    def auth_logout(request: Request):
        _destroy_session(_token_from_request(request))
        return {"ok": True}

    @router.get("/session")
    def auth_session(request: Request):
        session = _get_session(_token_from_request(request))
        if not session:
            return {"authenticated": False}
        return {
            "authenticated": True,
            "username": session["username"],
            "displayName": session["displayName"],
            "role": session["role"],
        }

    # Alias, falls Frontend /api/auth/me erwartet
    @router.get("/me")
    def auth_me(request: Request):
        return auth_session(request)

    @router.get("/users")
    def list_users(request: Request):
        session = _get_session(_token_from_request(request))
        if not session or session.get("role") != "admin":
            raise HTTPException(status_code=403, detail="Nur Administratoren.")
        users = _read_users(base_dir)
        return {
            "users": [
                {
                    "username": u.get("username"),
                    "displayName": u.get("displayName"),
                    "role": u.get("role"),
                }
                for u in users
            ]
        }

    @router.post("/users")
    def create_user(body: UserCreate, request: Request):
        session = _get_session(_token_from_request(request))
        if not session or session.get("role") != "admin":
            raise HTTPException(status_code=403, detail="Nur Administratoren.")
        if _find_user(base_dir, body.username):
            raise HTTPException(status_code=409, detail=f"Benutzer '{body.username}' existiert bereits.")
        if body.role not in ("admin", "user"):
            raise HTTPException(status_code=400, detail="Rolle muss 'admin' oder 'user' sein.")
        users = _read_users(base_dir)
        users.append(
            {
                "username": body.username.strip(),
                "password": _hash_password(body.password),
                "displayName": (body.displayName or body.username).strip(),
                "role": body.role,
            }
        )
        _write_users(base_dir, users)
        return {"ok": True, "message": f"Benutzer '{body.username}' erstellt."}

    return router
