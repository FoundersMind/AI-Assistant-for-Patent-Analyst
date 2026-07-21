"""
Lumenci Spark agent — executes structured workspace actions from LLM output or user intent.
"""

from __future__ import annotations

import json
import re
from typing import Any, Dict, List, Optional, Tuple

from django.db import transaction
from django.db.models import Max
from django.shortcuts import get_object_or_404
from django.utils import timezone

from .models import Case, ChatMessage, ClaimChart, ClaimChartRow, RowChange
from .strength_llm import sync_claim_chart_strengths, sync_one_row_strength

VALID_ACTION_TYPES = frozenset(
    {
        "accept_all",
        "reject_all",
        "accept_suggestions",
        "accept_new_rows",
        "reject_new_rows",
        "accept_row",
        "reject_row",
        "undo",
        "redo",
        "reassess_strengths",
        "update_row",
        "add_row",
        "add_empty_row",
        "delete_row",
        "clear_chat",
        "clear_history",
        "set_instructions",
        "rename_chart",
        "rename_case",
        "create_case",
        "highlight_row",
    }
)


def normalize_actions(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    out: List[Dict[str, Any]] = []
    for item in raw[:20]:
        if not isinstance(item, dict):
            continue
        t = (item.get("type") or "").strip().lower()
        if t not in VALID_ACTION_TYPES:
            continue
        out.append({**item, "type": t})
    return out


def _parse_new_rows_inline(payload: Dict[str, Any]) -> List[Dict[str, str]]:
    raw = payload.get("new_rows") or []
    if not isinstance(raw, list):
        return []
    out: List[Dict[str, str]] = []
    for item in raw[:8]:
        if not isinstance(item, dict):
            continue
        claim = str(item.get("claim") or "").strip()
        evidence = str(item.get("evidence") or "").strip()
        reasoning = str(item.get("reasoning") or "").strip()
        if not (claim or evidence or reasoning):
            continue
        out.append({"claim": claim, "evidence": evidence, "reasoning": reasoning})
    return out


def normalize_suggestions(raw: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw, list):
        return []
    filtered: List[Dict[str, Any]] = []
    for s in raw:
        if not isinstance(s, dict):
            continue
        field = s.get("field")
        if field not in ("claim", "evidence", "reasoning"):
            continue
        try:
            rid = int(s.get("row_id"))
        except Exception:
            continue
        filtered.append(
            {
                "row_id": rid,
                "field": field,
                "old_text": s.get("old_text") or "",
                "new_text": s.get("new_text") or "",
            }
        )
    return filtered


def parse_agent_payload(payload: Any) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]], List[Dict[str, Any]], bool]:
    """Returns suggestions, new_rows, actions, auto_apply from parsed JSON payload."""
    if not isinstance(payload, dict):
        return [], [], [], False
    suggestions = normalize_suggestions(payload.get("suggestions"))
    new_rows = _parse_new_rows_inline(payload)
    actions = normalize_actions(payload.get("actions"))
    auto_apply = bool(payload.get("auto_apply"))
    return suggestions, new_rows, actions, auto_apply


def _invalidate_redo_branch(ch: ClaimChart) -> None:
    RowChange.objects.filter(claim_chart=ch, is_undone=True).update(redo_invalidated=True)


def _apply_suggestions(ch: ClaimChart, suggestions: List[Dict[str, Any]]) -> int:
    if not suggestions:
        return 0
    _invalidate_redo_branch(ch)
    applied = 0
    with transaction.atomic():
        for suggestion in suggestions:
            row_id = int(suggestion.get("row_id") or 0)
            field = suggestion.get("field")
            new_text = suggestion.get("new_text") or ""
            if row_id <= 0 or field not in ("claim", "evidence", "reasoning"):
                continue
            row = ClaimChartRow.objects.filter(claim_chart=ch, row_index=row_id).first()
            if not row:
                continue
            old_text = (
                row.claim_text
                if field == "claim"
                else row.evidence_text
                if field == "evidence"
                else row.reasoning_text
            )
            RowChange.objects.create(
                claim_chart=ch,
                row_index=row_id,
                field=field,
                old_text=old_text,
                new_text=new_text,
            )
            if field == "claim":
                row.claim_text = new_text
            elif field == "evidence":
                row.evidence_text = new_text
            else:
                row.reasoning_text = new_text
            row.save(update_fields=["claim_text", "evidence_text", "reasoning_text"])
            sync_one_row_strength(row)
            applied += 1
    return applied


def _apply_new_rows(ch: ClaimChart, new_rows: List[Dict[str, str]]) -> int:
    if not new_rows:
        return 0
    _invalidate_redo_branch(ch)
    added = 0
    with transaction.atomic():
        for nr in new_rows:
            claim = (nr.get("claim") or "").strip()
            evidence = (nr.get("evidence") or "").strip()
            reasoning = (nr.get("reasoning") or "").strip()
            if not (claim or evidence or reasoning):
                continue
            next_idx = (ch.rows.aggregate(m=Max("row_index"))["m"] or 0) + 1
            snapshot = json.dumps(
                {
                    "claim": claim,
                    "evidence": evidence,
                    "reasoning": reasoning,
                    "origin": ClaimChartRow.RowOrigin.ADDED,
                },
                ensure_ascii=False,
            )
            RowChange.objects.create(
                claim_chart=ch,
                row_index=next_idx,
                field="add_row",
                old_text="",
                new_text=snapshot,
            )
            new_row = ClaimChartRow.objects.create(
                claim_chart=ch,
                row_index=next_idx,
                origin=ClaimChartRow.RowOrigin.ADDED,
                claim_text=claim,
                evidence_text=evidence,
                reasoning_text=reasoning,
            )
            sync_one_row_strength(new_row)
            added += 1
    return added


def _do_undo(ch: ClaimChart) -> Tuple[bool, str]:
    last = (
        RowChange.objects.filter(claim_chart=ch, is_undone=False)
        .order_by("-created_at", "-id")
        .first()
    )
    if not last:
        return False, "Nothing to undo"
    row_to_resync = None
    with transaction.atomic():
        if last.field == "add_row":
            ClaimChartRow.objects.filter(claim_chart=ch, row_index=last.row_index).delete()
        else:
            row = get_object_or_404(ClaimChartRow, claim_chart=ch, row_index=last.row_index)
            if last.field == "claim":
                row.claim_text = last.old_text
            elif last.field == "evidence":
                row.evidence_text = last.old_text
            else:
                row.reasoning_text = last.old_text
            row.save(update_fields=["claim_text", "evidence_text", "reasoning_text"])
            row_to_resync = row
        last.is_undone = True
        last.undone_at = timezone.now()
        last.save(update_fields=["is_undone", "undone_at"])
    if row_to_resync is not None:
        sync_one_row_strength(row_to_resync)
    return True, "Undid last change"


def _do_redo(ch: ClaimChart) -> Tuple[bool, str]:
    redo_op = (
        RowChange.objects.filter(claim_chart=ch, is_undone=True, redo_invalidated=False)
        .order_by("-undone_at", "-id")
        .first()
    )
    if not redo_op:
        return False, "Nothing to redo"
    row_to_resync = None
    with transaction.atomic():
        if redo_op.field == "add_row":
            try:
                payload = json.loads(redo_op.new_text or "{}")
            except json.JSONDecodeError:
                payload = {}
            new_row = ClaimChartRow.objects.create(
                claim_chart=ch,
                row_index=redo_op.row_index,
                origin=payload.get("origin") or ClaimChartRow.RowOrigin.ADDED,
                claim_text=str(payload.get("claim") or ""),
                evidence_text=str(payload.get("evidence") or ""),
                reasoning_text=str(payload.get("reasoning") or ""),
            )
            row_to_resync = new_row
        else:
            row = get_object_or_404(ClaimChartRow, claim_chart=ch, row_index=redo_op.row_index)
            if redo_op.field == "claim":
                row.claim_text = redo_op.new_text
            elif redo_op.field == "evidence":
                row.evidence_text = redo_op.new_text
            else:
                row.reasoning_text = redo_op.new_text
            row.save(update_fields=["claim_text", "evidence_text", "reasoning_text"])
            row_to_resync = row
        redo_op.is_undone = False
        redo_op.undone_at = None
        redo_op.save(update_fields=["is_undone", "undone_at"])
    if row_to_resync is not None:
        sync_one_row_strength(row_to_resync)
    return True, "Redid last change"


def execute_agent_actions(
    ch: ClaimChart,
    actions: List[Dict[str, Any]],
    *,
    pending_suggestions: Optional[List[Dict[str, Any]]] = None,
    pending_new_rows: Optional[List[Dict[str, str]]] = None,
    recovered_suggestions: Optional[List[Dict[str, Any]]] = None,
    recovered_new_rows: Optional[List[Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """
    Run agent actions against a claim chart. Returns execution report for API + UI sync.
    """
    pending_suggestions = list(pending_suggestions or [])
    pending_new_rows = list(pending_new_rows or [])
    recovered_suggestions = list(recovered_suggestions or [])
    recovered_new_rows = list(recovered_new_rows or [])

    executed: List[Dict[str, Any]] = []
    clear_pending = False
    highlight_row_id: Optional[int] = None
    client_hints: List[Dict[str, Any]] = []
    messages: List[str] = []

    for action in actions:
        t = action["type"]

        if t == "highlight_row":
            try:
                highlight_row_id = int(action.get("row_id") or 0)
            except (TypeError, ValueError):
                pass
            executed.append({"type": t, "row_id": highlight_row_id, "ok": True})
            continue

        if t == "reject_all":
            clear_pending = True
            n = len(pending_suggestions) + len(pending_new_rows) + len(recovered_suggestions) + len(recovered_new_rows)
            pending_suggestions = []
            pending_new_rows = []
            recovered_suggestions = []
            recovered_new_rows = []
            messages.append(f"Dismissed {n} pending item(s).")
            executed.append({"type": t, "applied": n, "ok": True})
            continue

        if t == "reject_new_rows":
            n = len(pending_new_rows) + len(recovered_new_rows)
            pending_new_rows = []
            recovered_new_rows = []
            messages.append(f"Dismissed {n} proposed row(s).")
            executed.append({"type": t, "applied": n, "ok": True})
            continue

        if t == "reject_row":
            try:
                row_id = int(action.get("row_id") or 0)
            except (TypeError, ValueError):
                executed.append({"type": t, "ok": False, "error": "row_id required"})
                continue
            field = action.get("field")
            before = len(pending_suggestions) + len(recovered_suggestions)
            pending_suggestions = [
                s
                for s in pending_suggestions
                if not (int(s.get("row_id") or 0) == row_id and (not field or s.get("field") == field))
            ]
            recovered_suggestions = [
                s
                for s in recovered_suggestions
                if not (int(s.get("row_id") or 0) == row_id and (not field or s.get("field") == field))
            ]
            removed = before - len(pending_suggestions) - len(recovered_suggestions)
            messages.append(f"Rejected {removed} suggestion(s) for row {row_id}.")
            executed.append({"type": t, "row_id": row_id, "field": field, "applied": removed, "ok": True})
            continue

        if t == "accept_all":
            pool = pending_suggestions or recovered_suggestions
            applied = _apply_suggestions(ch, pool)
            added = 0
            if pending_new_rows or recovered_new_rows:
                added = _apply_new_rows(ch, pending_new_rows or recovered_new_rows)
            clear_pending = True
            pending_suggestions = []
            pending_new_rows = []
            recovered_suggestions = []
            recovered_new_rows = []
            messages.append(f"Accepted {applied} edit(s)" + (f" and added {added} row(s)." if added else "."))
            executed.append({"type": t, "applied": applied + added, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "accept_suggestions":
            explicit = action.get("suggestions")
            if isinstance(explicit, list) and explicit:
                pool = explicit
            else:
                pool = pending_suggestions or recovered_suggestions
            applied = _apply_suggestions(ch, pool)
            if applied:
                clear_pending = True
                pending_suggestions = []
                recovered_suggestions = []
            messages.append(f"Accepted {applied} suggestion(s).")
            executed.append({"type": t, "applied": applied, "ok": applied > 0})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "accept_new_rows":
            pool = pending_new_rows or recovered_new_rows
            if action.get("rows") and isinstance(action.get("rows"), list):
                pool = _parse_new_rows_inline({"new_rows": action.get("rows")})
            added = _apply_new_rows(ch, pool)
            pending_new_rows = []
            recovered_new_rows = []
            messages.append(f"Added {added} row(s) to the chart.")
            executed.append({"type": t, "applied": added, "ok": added > 0})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "accept_row":
            try:
                row_id = int(action.get("row_id") or 0)
            except (TypeError, ValueError):
                executed.append({"type": t, "ok": False, "error": "row_id required"})
                continue
            field = action.get("field")
            pool = pending_suggestions + recovered_suggestions
            matching = [
                s
                for s in pool
                if int(s.get("row_id") or 0) == row_id and (not field or s.get("field") == field)
            ]
            applied = _apply_suggestions(ch, matching)
            pending_suggestions = [s for s in pending_suggestions if s not in matching]
            recovered_suggestions = [s for s in recovered_suggestions if s not in matching]
            highlight_row_id = row_id
            messages.append(f"Accepted {applied} change(s) on row {row_id}.")
            executed.append({"type": t, "row_id": row_id, "applied": applied, "ok": applied > 0})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "undo":
            ok, msg = _do_undo(ch)
            messages.append(msg)
            executed.append({"type": t, "ok": ok, "applied": 1 if ok else 0})
            clear_pending = True
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "redo":
            ok, msg = _do_redo(ch)
            messages.append(msg)
            executed.append({"type": t, "ok": ok, "applied": 1 if ok else 0})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "reassess_strengths":
            sync_claim_chart_strengths(ch)
            messages.append("Re-assessed strength for all rows.")
            executed.append({"type": t, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "update_row":
            try:
                row_id = int(action.get("row_id") or 0)
            except (TypeError, ValueError):
                executed.append({"type": t, "ok": False})
                continue
            field = action.get("field")
            text = str(action.get("text") or action.get("new_text") or "")
            if row_id <= 0 or field not in ("claim", "evidence", "reasoning"):
                executed.append({"type": t, "ok": False})
                continue
            row = ClaimChartRow.objects.filter(claim_chart=ch, row_index=row_id).first()
            if not row:
                executed.append({"type": t, "ok": False})
                continue
            _invalidate_redo_branch(ch)
            old = row.claim_text if field == "claim" else row.evidence_text if field == "evidence" else row.reasoning_text
            with transaction.atomic():
                RowChange.objects.create(
                    claim_chart=ch, row_index=row_id, field=field, old_text=old, new_text=text
                )
                if field == "claim":
                    row.claim_text = text
                elif field == "evidence":
                    row.evidence_text = text
                else:
                    row.reasoning_text = text
                row.save(update_fields=["claim_text", "evidence_text", "reasoning_text"])
            sync_one_row_strength(row)
            highlight_row_id = row_id
            messages.append(f"Updated row {row_id} {field}.")
            executed.append({"type": t, "row_id": row_id, "field": field, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "add_row":
            added = _apply_new_rows(
                ch,
                [
                    {
                        "claim": str(action.get("claim") or ""),
                        "evidence": str(action.get("evidence") or ""),
                        "reasoning": str(action.get("reasoning") or ""),
                    }
                ],
            )
            messages.append(f"Added {added} row(s).")
            executed.append({"type": t, "applied": added, "ok": added > 0})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "add_empty_row":
            next_idx = (ch.rows.aggregate(m=Max("row_index"))["m"] or 0) + 1
            ClaimChartRow.objects.create(
                claim_chart=ch,
                row_index=next_idx,
                origin=ClaimChartRow.RowOrigin.ADDED,
            )
            highlight_row_id = next_idx
            messages.append(f"Added empty row {next_idx}.")
            executed.append({"type": t, "row_id": next_idx, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "delete_row":
            try:
                row_id = int(action.get("row_id") or 0)
            except (TypeError, ValueError):
                executed.append({"type": t, "ok": False})
                continue
            with transaction.atomic():
                ClaimChartRow.objects.filter(claim_chart=ch, row_index=row_id).delete()
                RowChange.objects.filter(claim_chart=ch, row_index=row_id).delete()
            messages.append(f"Deleted row {row_id}.")
            executed.append({"type": t, "row_id": row_id, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "clear_chat":
            ChatMessage.objects.filter(claim_chart=ch).delete()
            messages.append("Chat cleared.")
            executed.append({"type": t, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "clear_history":
            RowChange.objects.filter(claim_chart=ch).delete()
            messages.append("Edit history cleared.")
            executed.append({"type": t, "ok": True})
            ch = ClaimChart.objects.select_related("case").get(pk=ch.pk)
            continue

        if t == "set_instructions":
            text = str(action.get("text") or action.get("instructions") or "")
            ch.system_instructions = text
            ch.save(update_fields=["system_instructions"])
            messages.append("Custom instructions updated.")
            executed.append({"type": t, "ok": True})
            continue

        if t == "rename_chart":
            name = str(action.get("name") or "").strip()
            if name:
                ch.name = name
                ch.save(update_fields=["name"])
                messages.append(f"Chart renamed to “{name}”.")
                executed.append({"type": t, "name": name, "ok": True})
            continue

        if t == "rename_case":
            name = str(action.get("name") or "").strip()
            if name and ch.case_id:
                case = ch.case
                case.name = name
                case.save(update_fields=["name"])
                messages.append(f"Matter renamed to “{name}”.")
                executed.append({"type": t, "name": name, "ok": True})
            continue

        if t == "create_case":
            name = str(action.get("name") or "New Case").strip() or "New Case"
            c = Case.objects.create(name=name)
            client_hints.append({"type": "create_case", "case_id": c.id, "name": c.name})
            messages.append(f"Created matter “{name}”.")
            executed.append({"type": t, "case_id": c.id, "name": name, "ok": True})
            continue

    return {
        "executed_actions": executed,
        "messages": messages,
        "clear_pending": clear_pending,
        "highlight_row_id": highlight_row_id,
        "client_hints": client_hints,
        "chart": ch,
        "remaining_suggestions": pending_suggestions + recovered_suggestions,
        "remaining_new_rows": pending_new_rows + recovered_new_rows,
    }


def infer_actions_from_user_message(msg: str) -> List[Dict[str, Any]]:
    """Map natural-language user messages to agent actions (pre-LLM fast path)."""
    ml = (msg or "").strip().lower()
    if not ml:
        return []

    if ml in ("undo", "undo last change"):
        return [{"type": "undo"}]
    if ml in ("redo", "redo last change"):
        return [{"type": "redo"}]

    row_m = re.search(
        r"\b(accept|reject|approve|dismiss)\b\s+(?:the\s+)?(?:row\s+|element\s+)?#?(\d+)(?:\s+[-:]?\s*(claim|evidence|reasoning))?",
        msg or "",
        re.IGNORECASE,
    )
    if row_m:
        verb = row_m.group(1).lower()
        row_id = int(row_m.group(2))
        field = (row_m.group(3) or "").lower() or None
        action_type = "accept_row" if verb in ("accept", "approve") else "reject_row"
        out: Dict[str, Any] = {"type": action_type, "row_id": row_id}
        if field in ("claim", "evidence", "reasoning"):
            out["field"] = field
        return [out]

    reject_triggers = (
        "reject all",
        "reject everything",
        "dismiss all",
        "decline all",
        "discard all",
        "reject them",
        "no thanks",
        "don't apply",
        "do not apply",
        "skip",
    )
    if ml in ("reject", "no", "nope", "decline", "dismiss") or any(t in ml for t in reject_triggers):
        return [{"type": "reject_all"}]

    accept_triggers = (
        "accept all",
        "accept everything",
        "accept them",
        "accept those",
        "apply all",
        "do it",
        "go ahead",
        "apply",
        "apply it",
        "apply that",
        "implement",
        "proceed",
        "commit",
        "looks good",
        "sounds good",
        "yes accept",
        "approve all",
    )
    if ml in ("accept", "yes", "ok", "okay", "approved", "approve") or any(t in ml for t in accept_triggers):
        return [{"type": "accept_all"}]

    if any(t in ml for t in ("accept new row", "accept proposed row", "add the row", "add proposed row", "insert new row")):
        return [{"type": "accept_new_rows"}]
    if any(t in ml for t in ("reject new row", "dismiss proposed", "discard new row")):
        return [{"type": "reject_new_rows"}]
    if any(t in ml for t in ("reassess", "re-assess", "refresh strength", "update strength", "strength badges")):
        return [{"type": "reassess_strengths"}]
    if any(t in ml for t in ("clear chat", "reset chat", "wipe chat")):
        return [{"type": "clear_chat"}]
    if any(t in ml for t in ("clear history", "clear log", "wipe history")):
        return [{"type": "clear_history"}]
    if ml.startswith("rename chart to ") or ml.startswith("rename this chart to "):
        name = msg.split(" to ", 1)[-1].strip().strip('"').strip("'")
        if name:
            return [{"type": "rename_chart", "name": name}]
    if ml.startswith("rename matter to ") or ml.startswith("rename case to "):
        name = msg.split(" to ", 1)[-1].strip().strip('"').strip("'")
        if name:
            return [{"type": "rename_case", "name": name}]
    if ml.startswith("create matter") or ml.startswith("create case") or ml.startswith("new matter"):
        name = "New Case"
        if " called " in ml:
            name = msg.split(" called ", 1)[-1].strip().strip('"').strip("'") or name
        elif " named " in ml:
            name = msg.split(" named ", 1)[-1].strip().strip('"').strip("'") or name
        return [{"type": "create_case", "name": name}]

    return []
