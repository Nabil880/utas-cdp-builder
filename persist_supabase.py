# persist_supabase.py
from __future__ import annotations
import json, secrets, datetime as _dt
from typing import Any, Dict, List, Optional

import streamlit as st
from supabase import create_client, Client

@st.cache_resource(show_spinner=False)
def _sb() -> Client:
    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        raise RuntimeError("Supabase secrets missing. Add SUPABASE_URL and SUPABASE_SERVICE_KEY to secrets.")
    return create_client(url, key)

def _db_enabled() -> bool:
    return bool(st.secrets.get("SUPABASE_URL")) and bool(st.secrets.get("SUPABASE_SERVICE_KEY"))

def _now_ts() -> int:
    return int(_dt.datetime.now(_dt.timezone.utc).timestamp())

def _normalize(x: str | None) -> str:
    return (x or "").strip().lower()

# ---- Bundle helpers ----
def _current_draft_bundle_dict() -> dict:
    # Prefer your existing build_bundle() if present
    try:
        from CDPapp import build_bundle  # type: ignore
        return json.loads(build_bundle())
    except Exception:
        pass
    draft = st.session_state.get("draft")
    if isinstance(draft, dict):
        return draft
    return {"_owner_uid": st.session_state.get("user_code","")}

# ---- Snapshots ----
def _persist_draft_snapshot(draft_id: str) -> None:
    if not _db_enabled():
        return
    data = _current_draft_bundle_dict()
    data["_owner_uid"] = st.session_state.get("user_code") or data.get("_owner_uid","")
    rec = {
        "draft_id": draft_id,
        "owner_uid": data.get("_owner_uid",""),
        "json": data,
        "updated_at": _dt.datetime.now(_dt.timezone.utc).isoformat()
    }
    _sb().table("draft_snapshots").upsert(rec, on_conflict="draft_id").execute()

def _load_snapshot_if_any(draft_id: str) -> Optional[dict]:
    if not _db_enabled():
        return None
    res = _sb().table("draft_snapshots").select("json").eq("draft_id", draft_id).execute()
    rows = res.data or []
    return rows[0].get("json") if rows else None

# ---- Tokens ----
def _issue_sign_token(payload: Dict[str, Any]) -> str:
    if not _db_enabled():
        tok = secrets.token_urlsafe(24)
        st.session_state.setdefault("_local_tokens", {})
        payload = dict(payload)
        payload["owner_uid"] = st.session_state.get("user_code","")
        payload["issued_at"] = _now_ts()
        st.session_state["_local_tokens"][tok] = {
            "payload_json": json.dumps(payload, ensure_ascii=False),
            "issued_at": payload["issued_at"],
            "used_at": None, "used_by": "", "note": ""
        }
        return tok

    tok = secrets.token_urlsafe(24)
    payload = dict(payload)
    payload["owner_uid"] = st.session_state.get("user_code","")
    payload["issued_at"] = _now_ts()
    _sb().table("tokens").insert({
        "token": tok,
        "payload_json": payload,
        "issued_at": payload["issued_at"],
        "used_at": None,
        "used_by": "",
        "note": ""
    }).execute()
    return tok

def _read_tokens() -> Dict[str, Dict[str, Any]]:
    if not _db_enabled():
        return st.session_state.get("_local_tokens", {})
    out: Dict[str, Dict[str, Any]] = {}
    res = _sb().table("tokens").select("token,payload_json,issued_at,used_at,used_by,note").execute()
    for row in (res.data or []):
        tok = row.get("token")
        if not tok:
            continue
        out[tok] = {
            "payload_json": json.dumps(row.get("payload_json", {}), ensure_ascii=False),
            "issued_at": row.get("issued_at"),
            "used_at": row.get("used_at"),
            "used_by": row.get("used_by",""),
            "note": row.get("note",""),
        }
    return out

def _mark_token_used(token: str):
    supabase.table("sign_tokens").update(
        {"used_at": int(time.time())}
    ).eq("token", token).execute()

# ---- “My tasks” / “Issued links” helpers ----
def _pending_sign_tasks_for_me() -> List[Dict[str, Any]]:
    me = st.session_state.get("user_profile") or {}
    me_name = _normalize(me.get("name"))
    me_email = _normalize(me.get("email"))
    rows: List[Dict[str, Any]] = []
    for tok, info in _read_tokens().items():
        if info.get("used_at"):
            continue
        try:
            p = json.loads(info.get("payload_json","{}"))
        except Exception:
            p = {}
        who = _normalize(p.get("name")) or _normalize(p.get("email"))
        if who and (who == me_name or who == me_email):
            rows.append({
                "token": tok,
                "row_type": p.get("row_type",""),
                "row_index": p.get("row_index",""),
                "draft_id": p.get("draft_id",""),
                "name": p.get("name",""),
                "email": p.get("email",""),
                "sections": p.get("sections",""),
                "issued_at": info.get("issued_at"),
            })
    return rows

def _compute_draft_status(draft_id: str) -> str:
    toks = _read_tokens()
    used = sum(1 for t in toks.values() if json.loads(t["payload_json"]).get("draft_id")==draft_id and t.get("used_at"))
    pending = sum(1 for t in toks.values() if json.loads(t["payload_json"]).get("draft_id")==draft_id and not t.get("used_at"))
    if pending and used:  return f"In progress ({used} signed, {pending} pending)"
    if pending:           return f"Awaiting signatures ({pending})"
    if used:              return f"Signed ({used})"
    return "No sign-offs issued"

def _my_issued_links():
    uid = (st.session_state.get("user_code") or "").strip()
    if not uid:
        return []

    # Pull your rows from Supabase (pseudo-code)
    rows = supabase.table("sign_tokens")\
                   .select("*")\
                   .eq("owner_uid", uid)\
                   .order("issued_at", desc=True)\
                   .execute().data

    out = []
    for r in rows:
        did = r["draft_id"]
        out.append({
            "token":        r["token"],
            "draft_id":     did,
            "row_type":     r["row_type"],         # "prepared" or "approved"
            "row_index":    int(r.get("row_index", 0)),
            "name":         r.get("name",""),
            "email":        r.get("email",""),
            "sections":     r.get("sections",""),
            "course_code":  r.get("course_code",""),
            "course_title": r.get("course_title",""),
            "academic_year":r.get("academic_year",""),
            "semester":     r.get("semester",""),
            "issued_at":    r.get("issued_at"),
            "used_at":      r.get("used_at"),      # <- include this even if None
            "status":       _compute_draft_status(did),
            "link":         _sign_link_for(r["token"]),
        })
    return out


# ---- Signature records ----
def _save_signature_record(*, draft_id: str, row_type: str, row_index: int,
                           signer_name: str, sig_path: str) -> None:
    supabase.table("sign_records").upsert({
        "draft_id": draft_id,
        "row_type": row_type,              # "prepared" / "approved"
        "row_index": row_index,            # 0 for approval
        "signer_name": signer_name,
        "sig_path": sig_path,
        "ts": int(time.time()),
    }).execute()


def _lookup_signature_record(draft_id: str, row_type: str, row_index: int) -> dict | None:
    res = supabase.table("sign_records")\
                  .select("*")\
                  .eq("draft_id", draft_id)\
                  .eq("row_type", row_type)\
                  .eq("row_index", row_index)\
                  .limit(1).execute().data
    return res[0] if res else None
