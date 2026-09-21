"""Data handling for the organogram app: load / save, validation, levels, tree helpers."""
from __future__ import annotations

import io
import json
from typing import Dict, List, Optional, Tuple

import pandas as pd

POS_COLS = [
    "position_id", "position_title", "name", "department", "reports_to",
    "status", "email", "phone", "sort_order",
]
EXTRA_COLS = ["position_id", "also_reports_to", "relationship"]
STATUS_OPTIONS = ["Filled", "Vacant"]
RELATIONSHIP_OPTIONS = ["Dotted line", "Functional", "Acting", "Project"]


# ----------------------------------------------------------------------------
# Creating / cleaning tables
# ----------------------------------------------------------------------------
def empty_positions() -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype="object") for c in POS_COLS})


def empty_extra() -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype="object") for c in EXTRA_COLS})


def _clean_text(v) -> str:
    """Return a stripped string; empty string for blanks/NaN."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def clean_positions(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only known columns, make everything text (except sort_order), drop empty rows."""
    df = df.copy()
    for c in POS_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[POS_COLS]
    for c in POS_COLS:
        if c != "sort_order":
            df[c] = df[c].map(_clean_text)
    df["sort_order"] = pd.to_numeric(df["sort_order"], errors="coerce")
    # a row with nothing in it is not a position
    keep = (df.drop(columns=["sort_order"]) != "").any(axis=1)
    df = df[keep].reset_index(drop=True)
    df["position_id"] = df["position_id"].str.upper()
    df["reports_to"] = df["reports_to"].str.upper()
    df["status"] = df["status"].str.capitalize()
    return df


def clean_extra(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in EXTRA_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[EXTRA_COLS]
    for c in EXTRA_COLS:
        df[c] = df[c].map(_clean_text)
    keep = (df != "").any(axis=1)
    df = df[keep].reset_index(drop=True)
    df["position_id"] = df["position_id"].str.upper()
    df["also_reports_to"] = df["also_reports_to"].str.upper()
    return df


# ----------------------------------------------------------------------------
# Load / save
# ----------------------------------------------------------------------------
def load_excel(file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Read the template (.xlsx). `file` may be a path or an uploaded file object."""
    xls = pd.ExcelFile(file, engine="openpyxl")
    sheets = {s.lower(): s for s in xls.sheet_names}
    if "positions" not in sheets:
        raise ValueError("The Excel file must contain a sheet named 'Positions'.")
    pos = pd.read_excel(xls, sheets["positions"], dtype=object)
    pos.columns = [str(c).strip() for c in pos.columns]
    extra = empty_extra()
    if "extra_reporting" in sheets:
        extra = pd.read_excel(xls, sheets["extra_reporting"], dtype=object)
        extra.columns = [str(c).strip() for c in extra.columns]
    missing = [c for c in ("position_id", "position_title") if c not in pos.columns]
    if missing:
        raise ValueError(f"Positions sheet is missing required column(s): {', '.join(missing)}")
    return clean_positions(pos), clean_extra(extra)


def to_excel_bytes(pos: pd.DataFrame, extra: pd.DataFrame) -> bytes:
    """Write the two tables to an Excel file (same layout as the template)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pos[POS_COLS].to_excel(writer, sheet_name="Positions", index=False)
        extra[EXTRA_COLS].to_excel(writer, sheet_name="Extra_Reporting", index=False)
        for name, widths in (
            ("Positions", [12, 30, 22, 20, 12, 10, 26, 16, 11]),
            ("Extra_Reporting", [14, 18, 20]),
        ):
            ws = writer.sheets[name]
            for i, w in enumerate(widths):
                ws.column_dimensions[chr(65 + i)].width = w
    return buf.getvalue()


def to_json_bytes(pos: pd.DataFrame, extra: pd.DataFrame) -> bytes:
    def records(df):
        out = []
        for rec in df.to_dict("records"):
            out.append({k: (None if (isinstance(v, float) and pd.isna(v)) else v) for k, v in rec.items()})
        return out

    return json.dumps({"positions": records(pos), "extra_reporting": records(extra)},
                      indent=2, ensure_ascii=False).encode("utf-8")


def load_json(file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw = file.read() if hasattr(file, "read") else open(file, "rb").read()
    data = json.loads(raw.decode("utf-8"))
    pos = pd.DataFrame(data.get("positions", []))
    extra = pd.DataFrame(data.get("extra_reporting", []))
    return clean_positions(pos), clean_extra(extra)


# ----------------------------------------------------------------------------
# Tree helpers
# ----------------------------------------------------------------------------
def positions_by_id(pos: pd.DataFrame) -> Dict[str, dict]:
    return {r["position_id"]: r for r in pos.to_dict("records") if r["position_id"]}


def children_map(pos: pd.DataFrame) -> Dict[str, List[str]]:
    """parent_id -> list of child ids, sorted by sort_order then file order.
    Top-level posts are stored under the key ''."""
    tmp = pos.reset_index(drop=True).copy()
    tmp["_order"] = tmp["sort_order"].fillna(1e9)
    tmp["_row"] = range(len(tmp))
    tmp = tmp.sort_values(["_order", "_row"])
    out: Dict[str, List[str]] = {}
    ids = set(tmp["position_id"])
    for _, r in tmp.iterrows():
        parent = r["reports_to"]
        if parent and parent not in ids:
            parent = ""  # broken link: show as top-level rather than losing it
        out.setdefault(parent, []).append(r["position_id"])
    return out


def compute_levels(pos: pd.DataFrame) -> Dict[str, int]:
    """Level 1 = no manager. Safe against loops (looping posts get level 0)."""
    by_id = positions_by_id(pos)
    levels: Dict[str, int] = {}
    for pid in by_id:
        chain, cur, seen = [], pid, set()
        while cur and cur in by_id and cur not in seen and cur not in levels:
            seen.add(cur)
            chain.append(cur)
            cur = by_id[cur]["reports_to"]
        if cur in seen:  # loop
            for c in chain:
                levels[c] = 0
            continue
        base = levels.get(cur, 0) if cur else 0
        for c in reversed(chain):
            base += 1
            levels[c] = base
    return levels


def descendants(pos: pd.DataFrame, root_id: str, include_root: bool = False) -> List[str]:
    cm = children_map(pos)
    out, stack, seen = [], [root_id], set()
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        if cur != root_id or include_root:
            out.append(cur)
        stack.extend(cm.get(cur, []))
    return out


def tree_rows(pos: pd.DataFrame) -> List[Tuple[str, int]]:
    """Depth-first list of (position_id, depth) with depth 0 for top-level posts."""
    cm = children_map(pos)
    rows: List[Tuple[str, int]] = []
    seen = set()

    def walk(pid: str, depth: int):
        if pid in seen:
            return
        seen.add(pid)
        rows.append((pid, depth))
        for c in cm.get(pid, []):
            walk(c, depth + 1)

    for top in cm.get("", []):
        walk(top, 0)
    return rows


def position_label(row: dict) -> str:
    title = row.get("position_title", "") or "(no title)"
    who = row.get("name", "")
    return f"{row['position_id']} - {title}" + (f" ({who})" if who else "")


# ----------------------------------------------------------------------------
# Adding / editing / deleting
# ----------------------------------------------------------------------------
def next_id(pos: pd.DataFrame, extra_taken: Optional[List[str]] = None) -> str:
    """P001, P002, ... one higher than the largest existing number."""
    nums = []
    for pid in list(pos["position_id"]) + list(extra_taken or []):
        digits = "".join(ch for ch in str(pid) if ch.isdigit())
        if str(pid).upper().startswith("P") and digits:
            nums.append(int(digits))
    return f"P{(max(nums) + 1 if nums else 1):03d}"


def next_sort_order(pos: pd.DataFrame, parent_id: str) -> int:
    sib = pd.to_numeric(pos[pos["reports_to"] == parent_id]["sort_order"], errors="coerce").dropna()
    return int(sib.max()) + 1 if len(sib) else 1


def add_position(pos, extra, *, title, name="", department="", reports_to="",
                 status=None, email="", phone="", also_reports_to=None, relationship=""):
    pid = next_id(pos)
    if status is None:
        status = "Filled" if name.strip() else "Vacant"
    row = {
        "position_id": pid, "position_title": title.strip(), "name": name.strip(),
        "department": department.strip(), "reports_to": reports_to, "status": status,
        "email": email.strip(), "phone": phone.strip(),
        "sort_order": next_sort_order(pos, reports_to),
    }
    pos = pd.concat([pos, pd.DataFrame([row])], ignore_index=True)
    pos["sort_order"] = pd.to_numeric(pos["sort_order"], errors="coerce")
    extra = set_extra_managers(extra, pid, also_reports_to or [], relationship)
    return pos, extra, pid


def update_position(pos, extra, pid, *, also_reports_to=None, relationship="", **fields):
    idx = pos.index[pos["position_id"] == pid]
    if len(idx) == 0:
        return pos, extra
    pos = pos.copy()
    for k, v in fields.items():
        if k in POS_COLS and k != "position_id":
            pos.loc[idx[0], k] = v.strip() if isinstance(v, str) else v
    if also_reports_to is not None:
        extra = set_extra_managers(extra, pid, also_reports_to, relationship)
    return pos, extra


def set_extra_managers(extra, pid, managers: List[str], relationship: str = "") -> pd.DataFrame:
    """Replace all extra-manager rows of `pid` with `managers` (keeps old relationship labels)."""
    old = {r["also_reports_to"]: r["relationship"] for r in extra.to_dict("records")
           if r["position_id"] == pid}
    rest = extra[extra["position_id"] != pid]
    new_rows = [{"position_id": pid, "also_reports_to": m,
                 "relationship": relationship or old.get(m, "")} for m in managers]
    if new_rows:
        rest = pd.concat([rest, pd.DataFrame(new_rows)], ignore_index=True)
    return rest.reset_index(drop=True)


def delete_position(pos, extra, pid, with_subtree: bool = False):
    """Delete a post. If it has sub-posts, either delete them too (with_subtree) or
    move them up to the deleted post's manager."""
    row = pos[pos["position_id"] == pid]
    if row.empty:
        return pos, extra
    parent = row.iloc[0]["reports_to"]
    if with_subtree:
        gone = set(descendants(pos, pid, include_root=True))
        pos = pos[~pos["position_id"].isin(gone)]
    else:
        pos = pos.copy()
        pos.loc[pos["reports_to"] == pid, "reports_to"] = parent
        pos = pos[pos["position_id"] != pid]
        gone = {pid}
    extra = extra[~extra["position_id"].isin(gone) & ~extra["also_reports_to"].isin(gone)]
    return pos.reset_index(drop=True), extra.reset_index(drop=True)


# ----------------------------------------------------------------------------
# Validation
# ----------------------------------------------------------------------------
def validate(pos: pd.DataFrame, extra: pd.DataFrame) -> Tuple[List[str], List[str]]:
    """Return (errors, warnings). Row numbers match the Excel sheet (header = row 1)."""
    errors: List[str] = []
    warnings: List[str] = []
    if pos.empty:
        return ["No positions entered yet."], warnings

    ids = list(pos["position_id"])
    id_set = {i for i in ids if i}

    for i, r in pos.reset_index(drop=True).iterrows():
        row = i + 2
        if not r["position_id"]:
            errors.append(f"Positions row {row}: position_id is blank.")
        if not r["position_title"]:
            errors.append(f"Positions row {row}: position_title is blank ({r['position_id'] or 'no id'}).")
        if r["reports_to"] and r["reports_to"] not in id_set:
            errors.append(f"Positions row {row}: reports_to '{r['reports_to']}' does not exist "
                          f"(for {r['position_id']}).")
        if r["reports_to"] and r["reports_to"] == r["position_id"]:
            errors.append(f"Positions row {row}: {r['position_id']} reports to itself.")
        if r["status"] and r["status"] not in STATUS_OPTIONS:
            warnings.append(f"Positions row {row}: status '{r['status']}' is not Filled/Vacant.")
        if r["status"] == "Vacant" and r["name"]:
            warnings.append(f"Positions row {row}: {r['position_id']} is Vacant but has a name.")

    seen = {}
    for i, pid in enumerate(ids):
        if pid in seen:
            errors.append(f"Positions row {i + 2}: duplicate position_id '{pid}' "
                          f"(also in row {seen[pid]}).")
        else:
            seen[pid] = i + 2

    # loops in the main reporting line
    by_id = positions_by_id(pos)
    looped = set()
    for pid in by_id:
        cur, path = pid, []
        while cur and cur in by_id and cur not in path:
            path.append(cur)
            cur = by_id[cur]["reports_to"]
        if cur in path:
            looped.update(path[path.index(cur):])
    if looped:
        errors.append("Reporting loop found among: " + ", ".join(sorted(looped)) + ".")

    if not any(r == "" for r in pos["reports_to"]):
        errors.append("No top-level (Level 1) position: at least one row must have a blank reports_to.")

    # extra reporting
    seen_pairs = set()
    for i, r in extra.reset_index(drop=True).iterrows():
        row = i + 2
        a, b = r["position_id"], r["also_reports_to"]
        if a not in id_set:
            errors.append(f"Extra_Reporting row {row}: position_id '{a}' does not exist.")
        if b not in id_set:
            errors.append(f"Extra_Reporting row {row}: also_reports_to '{b}' does not exist.")
        if a and a == b:
            errors.append(f"Extra_Reporting row {row}: {a} cannot report to itself.")
        if a in by_id and by_id[a]["reports_to"] == b:
            errors.append(f"Extra_Reporting row {row}: {b} is already the main manager of {a}.")
        if (a, b) in seen_pairs:
            warnings.append(f"Extra_Reporting row {row}: duplicate line {a} -> {b}.")
        seen_pairs.add((a, b))
        # an extra line that points at the post's own subordinate makes a cycle of authority
        if a in id_set and b in id_set and a != b and b in set(descendants(pos, a)):
            warnings.append(f"Extra_Reporting row {row}: {b} works under {a}, so this line creates a circular reporting.")
    return errors, warnings
