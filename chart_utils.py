"""Builds the organogram as Graphviz graphs (one graph per chart page)."""
from __future__ import annotations

import html
import textwrap
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import graphviz
import pandas as pd

from data_utils import children_map, compute_levels, descendants, positions_by_id

PALETTE = [
    "#DCE6F2", "#E2EFDA", "#FFF2CC", "#FCE4D6", "#EADCF4", "#DDEBF7",
    "#F8CBAD", "#D9F0EE", "#F4E1EC", "#EDEDED", "#E4DFEC", "#FFE5B4",
]
BORDER_COLOR = "#44546A"
VACANT_COLOR = "#C00000"
EXTRA_LINE_COLOR = "#2E75B6"
NO_DEPT = "(no department)"


@dataclass
class ChartOptions:
    rankdir: str = "TB"              # "TB" top-to-bottom or "LR" left-to-right
    line_style: str = "polyline"     # "polyline" (straight), "spline" (curved), "ortho" (elbow, experimental)
    color_by_department: bool = True
    show_id: bool = False
    show_contact: bool = False       # email / phone inside the boxes
    highlight_vacant: bool = True
    wrap: int = 24                   # characters per line inside a box
    fontname: str = "Helvetica"
    institution: str = ""


def department_colors(pos: pd.DataFrame) -> Dict[str, str]:
    """Same department always gets the same colour (alphabetical order)."""
    depts = sorted({d for d in pos["department"] if d})
    return {d: PALETTE[i % len(PALETTE)] for i, d in enumerate(depts)}


def _e(text) -> str:
    return html.escape(str(text), quote=True)


def _wrap_html(text: str, width: int) -> str:
    lines = textwrap.wrap(str(text), width=width) or [""]
    return "<BR/>".join(_e(l) for l in lines)


def _node_label(row: dict, notes: List[str], opts: ChartOptions) -> str:
    vacant = opts.highlight_vacant and row["status"] == "Vacant"
    rows = [f'<TR><TD><B>{_wrap_html(row["position_title"], opts.wrap)}</B></TD></TR>']
    if row["name"]:
        rows.append(f'<TR><TD>{_wrap_html(row["name"], opts.wrap)}</TD></TR>')
    elif vacant:
        rows.append(f'<TR><TD><FONT COLOR="{VACANT_COLOR}"><B>VACANT</B></FONT></TD></TR>')
    if row["department"]:
        rows.append(f'<TR><TD><FONT POINT-SIZE="9" COLOR="#595959">{_wrap_html(row["department"], opts.wrap + 4)}</FONT></TD></TR>')
    if opts.show_id:
        rows.append(f'<TR><TD><FONT POINT-SIZE="8" COLOR="#7F7F7F">{_e(row["position_id"])}</FONT></TD></TR>')
    if opts.show_contact:
        for key in ("email", "phone"):
            if row[key]:
                rows.append(f'<TR><TD><FONT POINT-SIZE="8" COLOR="#404040">{_e(row[key])}</FONT></TD></TR>')
    for n in notes:
        rows.append(f'<TR><TD><FONT POINT-SIZE="8" COLOR="{EXTRA_LINE_COLOR}"><I>{_wrap_html(n, opts.wrap + 6)}</I></FONT></TD></TR>')
    return ('<<TABLE BORDER="0" CELLBORDER="0" CELLSPACING="0" CELLPADDING="1">'
            + "".join(rows) + "</TABLE>>")


def _graph_title(title: str, opts: ChartOptions) -> str:
    """Institution name and page title only (no colour legend)."""
    head = title or ""
    if opts.institution:
        head = opts.institution + (f" - {title}" if title else "")
    if not head:
        return ""
    return ('<<TABLE BORDER="0" CELLBORDER="0" CELLSPACING="6"><TR><TD><FONT POINT-SIZE="22"><B>'
            + _e(head) + "</B></FONT></TD></TR></TABLE>>")


def build_graph(pos: pd.DataFrame, extra: pd.DataFrame, *, roots: Optional[List[str]] = None,
                max_depth: Optional[int] = None, title: str = "",
                opts: Optional[ChartOptions] = None,
                page_refs: Optional[Dict[str, int]] = None) -> graphviz.Digraph:
    """One chart page.

    roots      - posts to start from (None = all top-level posts)
    max_depth  - number of levels to draw counting the roots as level 1 (None = everything)
    page_refs  - {position_id: page number} used to write 'continues on page N'
    """
    opts = opts or ChartOptions()
    by_id = positions_by_id(pos)
    cm = children_map(pos)
    colors = department_colors(pos)
    top = roots if roots is not None else cm.get("", [])

    # ---- which posts are drawn ------------------------------------------------
    visible: Dict[str, int] = {}
    stack = [(r, 1) for r in reversed(top) if r in by_id]
    while stack:
        pid, d = stack.pop()
        if pid in visible:
            continue
        visible[pid] = d
        if max_depth is None or d < max_depth:
            stack.extend((c, d + 1) for c in reversed(cm.get(pid, [])))

    g = graphviz.Digraph(
        "organogram",
        graph_attr={
            "rankdir": opts.rankdir, "splines": opts.line_style, "nodesep": "0.3",
            "ranksep": "0.55", "pad": "0.3", "fontname": opts.fontname,
            "labelloc": "t", "newrank": "true", "bgcolor": "white",
        },
        node_attr={"shape": "box", "style": "rounded,filled", "fontname": opts.fontname,
                   "fontsize": "11", "margin": "0.12,0.07", "color": BORDER_COLOR,
                   "penwidth": "1.2"},
        edge_attr={"color": BORDER_COLOR, "arrowhead": "none", "penwidth": "1.1"},
    )

    extra_by_child: Dict[str, List[dict]] = {}
    for r in extra.to_dict("records"):
        extra_by_child.setdefault(r["position_id"], []).append(r)

    for pid in visible:  # dict keeps insertion order = tree order
        row = by_id[pid]
        notes: List[str] = []

        # branch root: say who it reports to
        if pid in top and roots is not None and row["reports_to"] in by_id and row["reports_to"] not in visible:
            m = by_id[row["reports_to"]]
            notes.append("reports to: " + m["position_title"] + (f" ({m['name']})" if m["name"] else ""))

        # extra managers outside this page become a note; inside become dashed edges
        for ex in extra_by_child.get(pid, []):
            mgr = by_id.get(ex["also_reports_to"])
            if not mgr:
                continue
            if ex["also_reports_to"] in visible:
                continue
            rel = f"{ex['relationship']}: " if ex["relationship"] else "also reports to: "
            notes.append(rel + mgr["position_title"] + (f" ({mgr['name']})" if mgr["name"] else ""))

        # people elsewhere who also report to this post (dotted line)
        subs = [by_id[r["position_id"]]["position_title"] for r in extra.to_dict("records")
                if r["also_reports_to"] == pid and r["position_id"] in by_id
                and r["position_id"] not in visible]
        if subs:
            shown_subs = ", ".join(subs[:3]) + (f" +{len(subs) - 3} more" if len(subs) > 3 else "")
            notes.append("also manages: " + shown_subs)

        # children that are not drawn
        hidden = [c for c in cm.get(pid, []) if c not in visible]
        if hidden:
            total = len(descendants(pos, pid))
            shown = sum(1 for d in descendants(pos, pid) if d in visible)
            n_hidden = total - shown
            if page_refs and pid in page_refs:
                notes.append(f"+{n_hidden} below - continues on page {page_refs[pid]}")
            else:
                notes.append(f"+{n_hidden} more below")

        vacant = opts.highlight_vacant and row["status"] == "Vacant"
        fill = colors.get(row["department"], "#FFFFFF") if opts.color_by_department else "#FFFFFF"
        if vacant:
            g.node(pid, _node_label(row, notes, opts), fillcolor="#FFFFFF" if not opts.color_by_department else fill,
                   color=VACANT_COLOR, style="rounded,filled,dashed", penwidth="1.6")
        else:
            g.node(pid, _node_label(row, notes, opts), fillcolor=fill)

    # ---- main (solid) lines ---------------------------------------------------
    for pid in visible:
        parent = by_id[pid]["reports_to"]
        if parent in visible:
            g.edge(parent, pid)

    # ---- extra (dashed) lines -------------------------------------------------
    for r in extra.to_dict("records"):
        a, b = r["position_id"], r["also_reports_to"]
        if a in visible and b in visible:
            g.edge(b, a, style="dashed", color=EXTRA_LINE_COLOR, constraint="false")

    label = _graph_title(title, opts)
    if label:
        g.graph_attr["label"] = label
    return g


# ----------------------------------------------------------------------------
# Multi-page layouts
# ----------------------------------------------------------------------------
def make_pages(pos: pd.DataFrame, extra: pd.DataFrame, *, mode: str = "full",
               root_id: Optional[str] = None, max_depth: Optional[int] = None,
               split_level: int = 2, opts: Optional[ChartOptions] = None) -> List[dict]:
    """Return [{'title': str, 'graph': Digraph, 'count': int}, ...]

    mode = 'full'   : whole institution on one page
           'branch' : one chosen post and everything below it
           'split'  : overview of levels 1..split_level, then one page per post at
                      split_level that has sub-posts (best for 50+ posts)
    """
    opts = opts or ChartOptions()
    by_id = positions_by_id(pos)
    if pos.empty:
        return []

    def count(g_roots, depth):
        cm = children_map(pos)
        seen, stack = set(), [(r, 1) for r in g_roots]
        while stack:
            pid, d = stack.pop()
            if pid in seen:
                continue
            seen.add(pid)
            if depth is None or d < depth:
                stack.extend((c, d + 1) for c in cm.get(pid, []))
        return len(seen)

    if mode == "branch" and root_id in by_id:
        r = by_id[root_id]
        title = "Branch: " + r["position_title"] + (f" ({r['name']})" if r["name"] else "")
        return [{"title": title, "count": count([root_id], max_depth),
                 "graph": build_graph(pos, extra, roots=[root_id], max_depth=max_depth,
                                      title=title, opts=opts)}]

    if mode == "split":
        levels = compute_levels(pos)
        cm = children_map(pos)
        branch_roots = [pid for pid in _tree_order(pos)
                        if levels.get(pid) == split_level and cm.get(pid)]
        page_refs = {pid: i + 2 for i, pid in enumerate(branch_roots)}
        pages = [{
            "title": f"Overview (levels 1-{split_level})",
            "count": count(cm.get("", []), split_level),
            "graph": build_graph(pos, extra, max_depth=split_level, opts=opts,
                                 title=f"Overview (levels 1-{split_level})", page_refs=page_refs),
        }]
        for pid in branch_roots:
            r = by_id[pid]
            title = f"{r['position_title']}" + (f" ({r['name']})" if r["name"] else "") + " - branch"
            pages.append({"title": title, "count": count([pid], max_depth),
                          "graph": build_graph(pos, extra, roots=[pid], max_depth=max_depth,
                                               title=title, opts=opts)})
        return pages

    cm = children_map(pos)
    return [{"title": "Organogram", "count": count(cm.get("", []), max_depth),
             "graph": build_graph(pos, extra, max_depth=max_depth, title="", opts=opts)}]


def _tree_order(pos: pd.DataFrame) -> List[str]:
    from data_utils import tree_rows
    return [pid for pid, _ in tree_rows(pos)]
