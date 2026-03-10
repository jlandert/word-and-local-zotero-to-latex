from __future__ import annotations

import json
from ast import literal_eval
from typing import Any
import base64

MENDELEY_PREFIX = "MENDELEY_CITATION_v3_"




def parse_mendeley_code(code: str) -> dict[str, Any] | None:
    
    """
    Parse the MENDELEY citation

    The payload is usually JSON. In some environments it may look like a Python
    literal; we therefore try JSON first and fall back to literal_eval.
    """
    if not isinstance(code, str):
        return None

    # Split off the prefix
    if not code.startswith(MENDELEY_PREFIX):
        return None
    
    b64_str = code.split("_v3_", 1)[1]  # get the Base64 part
    
    # Fix padding
    missing_padding = len(b64_str) % 4
    if missing_padding != 0:
        b64_str += "=" * (4 - missing_padding)
    
    # Decode Base64
    decoded_bytes = base64.b64decode(b64_str)
    raw = decoded_bytes.decode('utf-8')

    # Try JSON first (most common).
    try:
        parsed = json.loads(raw)
        return parsed if isinstance(parsed, dict) else None
    except Exception:
        pass

    # Fallback: best-effort Python literal parsing.
    try:
        parsed = literal_eval(raw)
        return parsed if isinstance(parsed, dict) else None
    except Exception:
        return None

