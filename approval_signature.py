import base64
import copy
import hmac
import hashlib
import json
from datetime import datetime, timezone
from typing import Any, Dict, Tuple


def _canonical_json_bytes(obj: Any) -> bytes:
    """
    Stable canonicalization:
    - sort_keys=True ensures consistent ordering
    - separators remove whitespace differences
    - ensure_ascii=False to keep UTF-8 stable
    """
    s = json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return s.encode("utf-8")


def sign_cdp_export(cdp_json: Dict[str, Any], *, approved_by: str, secret_key: str,
                    schema_version: str = "cdp-v1") -> Dict[str, Any]:
    """
    Returns a NEW dict with an embedded approval block + HMAC signature.
    Signature covers the JSON content EXCLUDING the approval block.
    """
    payload = copy.deepcopy(cdp_json)
    payload.pop("approval", None)

    msg = _canonical_json_bytes(payload)
    digest = hmac.new(secret_key.encode("utf-8"), msg, hashlib.sha256).digest()
    sig_b64 = base64.b64encode(digest).decode("ascii")

    signed = copy.deepcopy(payload)
    signed["approval"] = {
        "status": "approved",
        "approved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "approved_by": approved_by,
        "schema_version": schema_version,
        "signature_alg": "HMAC-SHA256",
        "signature": sig_b64,
    }
    return signed


def verify_approved_cdp_export(cdp_json: Dict[str, Any], *, secret_key: str) -> Tuple[bool, str]:
    """
    Validates approval block + signature.
    Returns (ok, message). Use compare_digest for timing-safe comparison.
    """
    approval = cdp_json.get("approval")
    if not isinstance(approval, dict):
        return False, "Missing approval block. Please upload an Approved CDP export."

    required = ["status", "approved_at", "approved_by", "schema_version", "signature_alg", "signature"]
    missing = [k for k in required if k not in approval]
    if missing:
        return False, f"Approval block is incomplete (missing: {', '.join(missing)})."

    if approval.get("status") != "approved":
        return False, "CDP export is not marked as approved."

    if approval.get("signature_alg") != "HMAC-SHA256":
        return False, "Unsupported signature algorithm."

    payload = copy.deepcopy(cdp_json)
    payload.pop("approval", None)

    msg = _canonical_json_bytes(payload)
    expected = hmac.new(secret_key.encode("utf-8"), msg, hashlib.sha256).digest()
    expected_b64 = base64.b64encode(expected).decode("ascii")

    if not hmac.compare_digest(expected_b64, str(approval.get("signature", ""))):
        return False, "Signature mismatch. This file was edited or is not an approved export."

    return True, "Approved CDP export verified."
