"""
LLM-based evidence-strength labels for claim-chart rows (Groq).
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List, Optional

VALID_STRENGTHS = frozenset({"strong", "weak", "missing"})

_STRENGTH_SYSTEM = """You are a senior patent-litigation analyst reviewing an infringement claim chart.

For EACH row you receive, read the three columns together:
- Patent claim element (what the patent requires)
- Accused product feature / evidence (what the accused product allegedly does or shows)
- Reasoning (how the evidence maps to the claim element)

Assign exactly one strength label per row using professional judgment—not keyword matching:
- "missing": No substantive evidence, evidence text empty or non-informative for mapping, or no real tie to the claim element.
- "weak": Some evidence but it is thin, vague, marketing-only, over-broad, incomplete, or the reasoning does not adequately connect the claim limitation to the evidence.
- "strong": Concrete, technical evidence with reasoning that clearly maps the claim limitation to a specific accused-product feature, behavior, or structure.

Return ONLY valid JSON (no markdown, no commentary) in this exact shape:
{"assessments":[{"row_id": <number>, "strength": "strong"|"weak"|"missing"}, ...]}

Include every row_id you were given, once each."""

_CHUNK = 12

_JSON_RE = re.compile(r"\{[\s\S]*\"assessments\"[\s\S]*\}")


def _minimal_fallback(evidence: str, reasoning: str) -> str:
    """Used only when Groq is unavailable or a row is missing from the model output."""
    if not (evidence or "").strip():
        return "missing"
    if not (reasoning or "").strip():
        return "weak"
    return "weak"


def _parse_assessments_json(raw: str) -> Dict[int, str]:
    out: Dict[int, str] = {}
    text = (raw or "").strip()
    if not text:
        return out
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        m = _JSON_RE.search(text)
        if not m:
            return out
        try:
            payload = json.loads(m.group(0))
        except json.JSONDecodeError:
            return out
    items = payload.get("assessments") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        return out
    for item in items:
        if not isinstance(item, dict):
            continue
        try:
            rid = int(item.get("row_id"))
        except Exception:
            continue
        s = item.get("strength")
        if isinstance(s, str) and s.strip().lower() in VALID_STRENGTHS:
            out[rid] = s.strip().lower()
    return out


def _groq_client():
    try:
        from django.conf import settings as dj_settings

        api_key = (getattr(dj_settings, "GROQ_API_KEY", "") or "").strip()
    except Exception:
        api_key = (os.getenv("GROQ_API_KEY") or "").strip().strip('"').strip("'")
    if not api_key:
        return None
    try:
        from groq import Groq

        return Groq(api_key=api_key)
    except Exception:
        return None


def assess_rows_with_groq(rows: List[Dict[str, Any]], model: Optional[str] = None) -> Dict[int, str]:
    """
    rows: list of dicts with keys row_id (int), claim, evidence, reasoning (str).
    Returns row_id -> strength for all input rows (fills gaps with minimal fallback).
    """
    if not rows:
        return {}

    by_row: Dict[int, Dict[str, Any]] = {}
    for r in rows:
        try:
            rid = int(r["row_id"])
        except Exception:
            continue
        by_row[rid] = r

    if not by_row:
        return {}

    client = _groq_client()
    try:
        from django.conf import settings as dj_settings

        default_model = getattr(dj_settings, "GROQ_MODEL", None) or "llama-3.3-70b-versatile"
    except Exception:
        default_model = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
    model_name = model or default_model
    merged: Dict[int, str] = {}

    if client is None:
        for rid, r in by_row.items():
            merged[rid] = _minimal_fallback(str(r.get("evidence") or ""), str(r.get("reasoning") or ""))
        return merged

    keys = list(by_row.keys())
    for i in range(0, len(keys), _CHUNK):
        chunk_ids = keys[i : i + _CHUNK]
        payload_rows = []
        for rid in chunk_ids:
            r = by_row[rid]
            payload_rows.append(
                {
                    "row_id": rid,
                    "claim": str(r.get("claim") or "")[:8000],
                    "evidence": str(r.get("evidence") or "")[:8000],
                    "reasoning": str(r.get("reasoning") or "")[:8000],
                }
            )
        user_msg = json.dumps({"rows": payload_rows}, ensure_ascii=False)
        try:
            resp = client.chat.completions.create(
                model=model_name,
                temperature=0.1,
                max_tokens=800,
                messages=[
                    {"role": "system", "content": _STRENGTH_SYSTEM},
                    {"role": "user", "content": user_msg},
                ],
            )
            raw = (resp.choices[0].message.content if getattr(resp, "choices", None) else "") or ""
            parsed = _parse_assessments_json(raw)
            for rid in chunk_ids:
                s = parsed.get(rid)
                if s in VALID_STRENGTHS:
                    merged[rid] = s
                else:
                    r = by_row[rid]
                    merged[rid] = _minimal_fallback(str(r.get("evidence") or ""), str(r.get("reasoning") or ""))
        except Exception:
            for rid in chunk_ids:
                r = by_row[rid]
                merged[rid] = _minimal_fallback(str(r.get("evidence") or ""), str(r.get("reasoning") or ""))

    return merged


def sync_claim_chart_strengths(ch) -> None:
    """Persist LLM strengths for all rows on chart. `ch` is ClaimChart."""
    from .models import ClaimChartRow

    rows = list(ClaimChartRow.objects.filter(claim_chart=ch).order_by("row_index"))
    if not rows:
        return
    payload = [
        {
            "row_id": r.row_index,
            "claim": r.claim_text,
            "evidence": r.evidence_text,
            "reasoning": r.reasoning_text,
        }
        for r in rows
    ]
    by_id = assess_rows_with_groq(payload)
    updated: List = []
    for r in rows:
        s = by_id.get(r.row_index)
        if s in VALID_STRENGTHS and r.strength != s:
            r.strength = s
            updated.append(r)
    if updated:
        ClaimChartRow.objects.bulk_update(updated, ["strength"], batch_size=50)


def sync_one_row_strength(row) -> Dict[str, Any]:
    """Re-assess a single ClaimChartRow after edit. Returns before/after strength metadata."""
    from .models import ClaimChartRow

    previous = (row.strength or ClaimChartRow.Strength.WEAK).strip().lower()
    if previous not in VALID_STRENGTHS:
        previous = "weak"

    by_id = assess_rows_with_groq(
        [
            {
                "row_id": row.row_index,
                "claim": row.claim_text,
                "evidence": row.evidence_text,
                "reasoning": row.reasoning_text,
            }
        ]
    )
    s = by_id.get(row.row_index)
    new = s if s in VALID_STRENGTHS else previous
    if s in VALID_STRENGTHS:
        ClaimChartRow.objects.filter(pk=row.pk).update(strength=s)
        row.strength = s
    return {
        "row_id": row.row_index,
        "previous_strength": previous,
        "new_strength": new,
        "unchanged": previous == new,
    }


_STRENGTH_RANK = {"missing": 0, "weak": 1, "strong": 2}


def format_strength_judgment(judgment: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Build a user-facing note when strength is reassessed after an edit."""
    if not isinstance(judgment, dict):
        return None
    try:
        row_id = int(judgment.get("row_id"))
    except (TypeError, ValueError):
        return None
    prev = str(judgment.get("previous_strength") or "weak").lower()
    new = str(judgment.get("new_strength") or "weak").lower()
    if prev not in VALID_STRENGTHS:
        prev = "weak"
    if new not in VALID_STRENGTHS:
        new = "weak"

    out = {
        "row_id": row_id,
        "previous_strength": prev,
        "new_strength": new,
        "unchanged": prev == new,
    }

    if prev == new:
        if new == "weak":
            out["message"] = (
                f"Row {row_id}: your edit was saved, but strength stays Weak. "
                "That is AI judgment—the evidence may still be thin or not clearly mapped to the claim. "
                "Add technical detail or override strength manually if you disagree."
            )
            return out
        if new == "missing":
            out["message"] = (
                f"Row {row_id}: your edit was saved, but strength stays Missing. "
                "AI judgment: substantive evidence for this claim element is still lacking."
            )
            return out
        return None

    if new == "strong" and _STRENGTH_RANK.get(prev, 0) < _STRENGTH_RANK["strong"]:
        out["improved"] = True
        out["message"] = f"Row {row_id}: strength upgraded to Strong after your edit."
        return out

    if _STRENGTH_RANK.get(new, 0) < _STRENGTH_RANK.get(prev, 0):
        out["improved"] = False
        out["message"] = f"Row {row_id}: edit saved; strength reassessed as {new.title()}."
        return out
    return None


def dedupe_strength_judgments(judgments: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Keep the latest judgment per row; drop rows that need no user message."""
    by_row: Dict[int, Dict[str, Any]] = {}
    for raw in judgments or []:
        formatted = format_strength_judgment(raw)
        if formatted:
            by_row[int(formatted["row_id"])] = formatted
    return list(by_row.values())
