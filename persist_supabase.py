# persist_supabase.py
# Single-responsibility persistence layer for Streamlit + Supabase
# Implements exactly the helpers your app imports:
#   persist_draft_snapshot, load_snapshot_if_any, issue_sign_token,
#   pending_sign_tasks_for_me, my_issued_links, mark_token_used,
#   save_signature_record

from __future__ import annotations
import time, json, secrets
from typing import Any, Dict, List, Optional

import streamlit as st
from supabase import create_client, Client

# ---------- Client (cached) ----------
_client: Optional[Client] = None

def _supabase() -> Client:
    global _client
    if _client is not None:
        return _client

    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_SERVICE_KEY")  # use service_role or make RLS policies
    if not url or not key:
        raise RuntimeError("Supabase secrets missing: SUPABASE_URL / SUPABASE_SERVICE_KEY")

    _client = create_client(url, key)
    return _client

def _now() -> int:
    return int(time.time())

def _norm(s: str | None) -> str:
    return (s or "").strip().lower()

# ---------- DRAFT SNAPSHOTS ----------
def persist_draft_snapshot(bundle: Dict[str, Any]) -> None:
    """
    Upsert the full CDP bundle (JSON) for draft_id.
    Your app gives this the same dict that powers the Sidebar JSON download.
    """
    sb = _supabase()
    did = bundle.get("_draft_id") or bundle.get("draft_id") or bundle.get("_di")
    if not did:
        # fall back: some versions keep draft id under session but the caller knows it;
        # if none found, do nothing (avoid crashing UI)
        return

    row = {
        "draft_id": str(did),
        "owner_uid": bundle.get("_owner_uid") or st.session_state.get("user_code", ""),
        "json": json.dumps(bundle, ensure_ascii=False),
        "updated_at": _now(),
    }

    # on_conflict ensures idempotent writes per draft_id
    sb.table("draft_snapshots").upsert(row, on_conflict="draft_id").execute()

def load_snapshot_if_any(draft_id: str) -> Optional[Dict[str, Any]]:
    """
    Returns the last saved bundle dict from Supabase if present; else None.
    """
    sb = _supabase()
    res = sb.table("draft_snapshots").select("*").eq("draft_id", draft_id).limit(1).execute()
    if not res.data:
        return None
    try:
        return json.loads(res.data[0]["json"])
    except Exception:
        return None

# ---------- SIGN TOKENS ----------
def issue_sign_token(payload: Dict[str, Any]) -> str:
    """
    Creates a token row in sign_tokens. Returns the token string.
    payload must include draft_id, row_type, row_index, name, email, sections, and some course metadata.
    """
    sb = _supabase()
    tok = secrets.token_urlsafe(24)

    payload = dict(payload)  # copy
    payload["issued_at"] = _now()
    payload["owner_uid"] = st.session_state.get("user_code", "")

    row = {
        "token": tok,
        "payload_json": json.dumps(payload, ensure_ascii=False),
        "issued_at": payload["issued_at"],
        "used_at": None,
        "used_by": "",
        "owner_uid": payload["owner_uid"],
        "note": "",
    }

    sb.table("sign_tokens").insert(row).execute()
    return tok

def mark_token_used(token: str, used_by: str = "") -> None:
    """
    Marks token as used; safe if token missing.
    """
    sb = _supabase()
    sb.table("sign_tokens").update({
        "used_at": _now(),
        "used_by": used_by or "",
    }).eq("token", token).execute()

# ---------- PENDING / ISSUED LISTS ----------
def _decode_payload_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for r in rows:
        try:
            p = json.loads(r.get("payload_json") or "{}")
        except Exception:
            p = {}
        out.append({
            "token": r.get("token", ""),
            "draft_id": p.get("draft_id", ""),
            "row_type": p.get("row_type", ""),
            "row_index": p.get("row_index", 0),
            "name": p.get("name", ""),
            "email": p.get("email", ""),
            "sections": p.get("sections", ""),
            "course_code": p.get("course_code", ""),
            "course_title": p.get("course_title", ""),
            "academic_year": p.get("academic_year", ""),
            "semester": p.get("semester", ""),
            "issued_at": r.get("issued_at"),
            "used_at": r.get("used_at"),      # <- keep key present (None if not used)
            "used_by": r.get("used_by", ""),
        })
    return out

def pending_sign_tasks_for_me() -> List[Dict[str, Any]]:
    """
    Tokens addressed to me (by email OR name), unused.
    Only call this after login (user profile available).
    """
    sb = _supabase()
    prof = st.session_state.get("user_profile") or {}
    me_email = _norm(prof.get("email"))
    me_name  = _norm(prof.get("name"))

    if not me_email and not me_name:
        return []

    # Pull latest, then filter in Python for name/email match to avoid complex text filters
    res = sb.table("sign_tokens").select("*").is_("used_at", None).order("issued_at", desc=True).execute()
    if not res.data:
        return []

    rows = _decode_payload_rows(res.data)

    out = []
    for r in rows:
        if _norm(r.get("email")) == me_email:
            out.append(r)
            continue
        # name fallback match
        if me_name and _norm(r.get("name")) == me_name:
            out.append(r)

    return out

def my_issued_links() -> List[Dict[str, Any]]:
    """
    Tokens I have issued (owner_uid == me), newest first.
    """
    sb = _supabase()
    uid = st.session_state.get("user_code", "")
    if not uid:
        return []

    res = sb.table("sign_tokens").select("*").eq("owner_uid", uid).order("issued_at", desc=True).execute()
    return _decode_payload_rows(res.data or [])

# ---------- SIGNATURE RECORDS ----------
def save_signature_record(draft_id: str, row_type: str, row_index: int,
                          signer_name: str, sig_path: str, email: str = "") -> None:
    """
    Stores who signed what, and where the signature image lives.
    """
    sb = _supabase()
    row = {
        "draft_id": draft_id,
        "row_type": row_type,           # 'prepared' | 'approved'
        "row_index": int(row_index),    # approvals use 0
        "name": signer_name,
        "email": email or "",
        "signature_path": sig_path,
        "ts": _now(),
    }
    sb.table("sign_records").insert(row).execute()
