from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime
from functools import wraps
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory, session
from werkzeug.security import check_password_hash, generate_password_hash


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "rsm_muestras.db"
DEFAULT_ADMIN_USER = os.environ.get("RSM_ADMIN_USER", "admin")
DEFAULT_ADMIN_PASSWORD = os.environ.get("RSM_ADMIN_PASSWORD", "C3l3st32018*")
LICENSE_PRODUCT_CODE = "RSM_MUESTRAS"
LICENSE_SECRET_MARK = "RSM_MUESTRAS_2026_LOCAL"

app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")
app.config["SECRET_KEY"] = os.environ.get("RSM_SECRET_KEY", "rsm_muestras_secret_2026")


def get_db() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def now_iso() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def hash_seed(value: str) -> int:
    text = str(value)
    hashed = 2166136261
    for char in text:
        hashed ^= ord(char)
        hashed += (hashed << 1) + (hashed << 4) + (hashed << 7) + (hashed << 8) + (hashed << 24)
    return abs(hashed & 0xFFFFFFFF)


def build_license_key(product: str, company: str, issued_to: str, expires_at: str) -> str:
    normalized_base = "|".join([
        product,
        company.upper(),
        issued_to.upper(),
        expires_at,
        LICENSE_SECRET_MARK
    ])
    hash_a = f"{hash_seed(normalized_base):08X}"
    hash_b = f"{hash_seed(f'RSM|{normalized_base}|CONTROL'):08X}"
    return f"{hash_a[:4]}-{hash_a[4:8]}-{hash_b[:4]}-{hash_b[4:8]}"


def normalize_license_payload(raw: dict) -> dict:
    return {
        "product": str(raw.get("product") or raw.get("producto") or "").strip().upper(),
        "company": str(raw.get("company") or raw.get("empresa") or "").strip(),
        "issued_to": str(raw.get("issuedTo") or raw.get("titular") or raw.get("usuario") or "").strip(),
        "expires_at": str(raw.get("expiresAt") or raw.get("vence") or raw.get("fechaVencimiento") or "").strip(),
        "license_key": str(raw.get("licenseKey") or raw.get("licencia") or raw.get("clave") or "").strip().upper(),
    }


def validate_license_payload(raw: dict) -> dict:
    license_data = normalize_license_payload(raw)

    if not all(license_data.values()):
        raise ValueError("La licencia está incompleta. Verifica producto, empresa, titular, vencimiento y clave.")

    if license_data["product"] != LICENSE_PRODUCT_CODE:
        raise ValueError("La licencia no corresponde a esta aplicación.")

    try:
        expires_at = datetime.strptime(license_data["expires_at"], "%Y-%m-%d").date()
    except ValueError as error:
        raise ValueError("La fecha de vencimiento debe estar en formato AAAA-MM-DD.") from error

    if expires_at.isoformat() < datetime.utcnow().date().isoformat():
        raise ValueError(f"La licencia venció el {license_data['expires_at']}.")

    expected_key = build_license_key(
        license_data["product"],
        license_data["company"],
        license_data["issued_to"],
        license_data["expires_at"]
    )
    if license_data["license_key"] != expected_key:
        raise ValueError("La clave de licencia no es válida para los datos suministrados.")

    return license_data


def require_admin(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("admin_user"):
            return jsonify({"ok": False, "message": "Sesión no autorizada."}), 401
        return view(*args, **kwargs)

    return wrapped


def init_db() -> None:
    with get_db() as db:
        db.execute(
            """
            CREATE TABLE IF NOT EXISTS admins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        db.execute(
            """
            CREATE TABLE IF NOT EXISTS licenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product TEXT NOT NULL,
                company TEXT NOT NULL,
                issued_to TEXT NOT NULL,
                expires_at TEXT NOT NULL,
                license_key TEXT NOT NULL,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )

        existing_admin = db.execute(
            "SELECT id FROM admins WHERE username = ?",
            (DEFAULT_ADMIN_USER,)
        ).fetchone()

        if not existing_admin:
            timestamp = now_iso()
            db.execute(
                """
                INSERT INTO admins (username, password_hash, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                """,
                (
                    DEFAULT_ADMIN_USER,
                    generate_password_hash(DEFAULT_ADMIN_PASSWORD),
                    timestamp,
                    timestamp
                )
            )
        db.commit()


@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/generador-licencias")
def admin_generator():
    return send_from_directory(BASE_DIR, "generador-licencias.html")


@app.route("/api/admin/session", methods=["GET"])
def admin_session():
    if not session.get("admin_user"):
        return jsonify({"ok": True, "authenticated": False})
    return jsonify({
        "ok": True,
        "authenticated": True,
        "user": session["admin_user"]
    })


@app.route("/api/admin/login", methods=["POST"])
def admin_login():
    data = request.get_json(silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not username or not password:
        return jsonify({"ok": False, "message": "Ingresa usuario y contraseña."}), 400

    with get_db() as db:
        admin = db.execute(
            "SELECT username, password_hash FROM admins WHERE username = ?",
            (username,)
        ).fetchone()

    if not admin or not check_password_hash(admin["password_hash"], password):
        return jsonify({"ok": False, "message": "Usuario o contraseña incorrectos."}), 401

    session["admin_user"] = admin["username"]
    return jsonify({
        "ok": True,
        "message": "Sesión iniciada correctamente.",
        "user": admin["username"]
    })


@app.route("/api/admin/logout", methods=["POST"])
def admin_logout():
    session.clear()
    return jsonify({"ok": True, "message": "Sesión cerrada correctamente."})


@app.route("/api/admin/change-password", methods=["POST"])
@require_admin
def admin_change_password():
    data = request.get_json(silent=True) or {}
    current_password = str(data.get("currentPassword", "")).strip()
    new_password = str(data.get("newPassword", "")).strip()
    confirm_password = str(data.get("confirmPassword", "")).strip()
    username = session["admin_user"]

    if not current_password or not new_password or not confirm_password:
        return jsonify({"ok": False, "message": "Completa clave actual, nueva clave y confirmación."}), 400

    if len(new_password) < 8:
        return jsonify({"ok": False, "message": "La nueva clave debe tener al menos 8 caracteres."}), 400

    if new_password != confirm_password:
        return jsonify({"ok": False, "message": "La confirmación no coincide con la nueva clave."}), 400

    with get_db() as db:
        admin = db.execute(
            "SELECT password_hash FROM admins WHERE username = ?",
            (username,)
        ).fetchone()

        if not admin or not check_password_hash(admin["password_hash"], current_password):
            return jsonify({"ok": False, "message": "La clave actual no es correcta."}), 401

        db.execute(
            "UPDATE admins SET password_hash = ?, updated_at = ? WHERE username = ?",
            (generate_password_hash(new_password), now_iso(), username)
        )
        db.commit()

    return jsonify({"ok": True, "message": "La clave del administrador se actualizó correctamente."})


@app.route("/api/licenses/generate", methods=["POST"])
@require_admin
def generate_license():
    data = request.get_json(silent=True) or {}
    company = str(data.get("company", "")).strip()
    issued_to = str(data.get("issuedTo", "")).strip()
    expires_at = str(data.get("expiresAt", "")).strip()

    if not company or not issued_to or not expires_at:
        return jsonify({"ok": False, "message": "Completa empresa, titular y fecha de vencimiento."}), 400

    try:
        datetime.strptime(expires_at, "%Y-%m-%d")
    except ValueError:
        return jsonify({"ok": False, "message": "La fecha de vencimiento debe estar en formato AAAA-MM-DD."}), 400

    if expires_at < datetime.utcnow().date().isoformat():
        return jsonify({"ok": False, "message": "La fecha de vencimiento no puede estar en el pasado."}), 400

    license_key = build_license_key(LICENSE_PRODUCT_CODE, company, issued_to, expires_at)
    timestamp = now_iso()
    created_by = session["admin_user"]

    with get_db() as db:
        db.execute(
            """
            INSERT INTO licenses (product, company, issued_to, expires_at, license_key, created_by, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                LICENSE_PRODUCT_CODE,
                company,
                issued_to,
                expires_at,
                license_key,
                created_by,
                timestamp
            )
        )
        db.commit()

    payload = {
        "producto": LICENSE_PRODUCT_CODE,
        "empresa": company,
        "titular": issued_to,
        "vence": expires_at,
        "licencia": license_key
    }
    return jsonify({
        "ok": True,
        "message": "Licencia generada correctamente.",
        "license": payload
    })


@app.route("/api/licenses/validate", methods=["POST"])
def validate_license():
    data = request.get_json(silent=True) or {}

    try:
        license_data = validate_license_payload(data)
        return jsonify({
            "ok": True,
            "message": "Licencia válida.",
            "license": {
                "product": license_data["product"],
                "company": license_data["company"],
                "issuedTo": license_data["issued_to"],
                "expiresAt": license_data["expires_at"],
                "licenseKey": license_data["license_key"]
            }
        })
    except ValueError as error:
        return jsonify({"ok": False, "message": str(error)}), 400


@app.route("/<path:filename>")
def static_files(filename: str):
    return send_from_directory(BASE_DIR, filename)


if __name__ == "__main__":
    init_db()
    app.run(host="127.0.0.1", port=8000, debug=True)
