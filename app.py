"""Organogram Builder - Streamlit app.

Run locally:   streamlit run app.py
"""
from __future__ import annotations

import datetime
import hashlib
import re
from pathlib import Path

import pandas as pd
import streamlit as st

import data_utils as du
from chart_utils import ChartOptions, make_pages
from export_utils import PAPERS, build_word, graph_to_png, pages_to_pdf

APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "sample_data.xlsx"

st.set_page_config(page_title="Organogram Builder", page_icon="🏛️", layout="wide")


# =============================================================================
# State helpers
# =============================================================================
def init_state():
    ss = st.session_state
    ss.setdefault("pos", du.empty_positions())
    ss.setdefault("extra", du.empty_extra())
    ss.setdefault("form", None)        # {"mode": "add"|"edit"|"delete", "pid": str|None, "parent": str}
    ss.setdefault("nonce", 0)          # changes a form's key so it starts empty again
    ss.setdefault("l1_rows", 3)
    ss.setdefault("exports", None)


def set_data(pos, extra):
    st.session_state.pos = pos
    st.session_state.extra = extra
    st.session_state.exports = None


def open_form(mode: str, pid: str | None = None, parent: str = ""):
    st.session_state.form = {"mode": mode, "pid": pid, "parent": parent}


def close_form():
    st.session_state.form = None


def more_rows():
    st.session_state.l1_rows += 3


def md_escape(text: str) -> str:
    return re.sub(r"([\\`*_{}\[\]#<>|~$])", r"\\\1", str(text))


def slug(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_") or "organogram"


def label_map(pos: pd.DataFrame) -> dict:
    return {r["position_id"]: du.position_label(r) for r in pos.to_dict("records")}


# =============================================================================
# Sidebar
# =============================================================================
def sidebar():
    ss = st.session_state
    st.sidebar.title("🏛️ Organogram Builder")
    st.sidebar.text_input("Institution name (shown on chart)", key="institution")

    st.sidebar.warning(
        "**Your data is not stored on the server.** When you close the page or the app "
        "restarts, it is gone. Use *Save your data* below and upload the file next time."
    )

    st.sidebar.subheader("Open data")
    up = st.sidebar.file_uploader("Excel (.xlsx) or JSON (.json)", type=["xlsx", "json"])
    if up is not None and st.sidebar.button("Load this file", type="primary"):
        try:
            if up.name.lower().endswith(".json"):
                pos, extra = du.load_json(up)
            else:
                pos, extra = du.load_excel(up)
            set_data(pos, extra)
            close_form()
            st.session_state.flash = f"Loaded {len(pos)} positions."
            st.rerun()
        except Exception as exc:  # show a friendly message instead of a traceback
            st.sidebar.error(f"Could not read the file: {exc}")

    if TEMPLATE_PATH.exists():
        if st.sidebar.button("Load sample data"):
            pos, extra = du.load_excel(TEMPLATE_PATH)
            set_data(pos, extra)
            close_form()
            st.rerun()
        st.sidebar.download_button(
            "⬇ Download blank Excel template", TEMPLATE_PATH.read_bytes(),
            file_name="organogram_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.sidebar.subheader("Save your data")
    if ss.pos.empty:
        st.sidebar.caption("Nothing to save yet.")
    else:
        stamp = datetime.date.today().isoformat()
        st.sidebar.download_button(
            "⬇ Save as Excel", du.to_excel_bytes(ss.pos, ss.extra),
            file_name=f"organogram_data_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.sidebar.download_button(
            "⬇ Save as JSON", du.to_json_bytes(ss.pos, ss.extra),
            file_name=f"organogram_data_{stamp}.json", mime="application/json",
        )

    with st.sidebar.expander("Start again"):
        sure = st.checkbox("Yes, delete everything I entered")
        if st.button("Clear all data", disabled=not sure):
            set_data(du.empty_positions(), du.empty_extra())
            close_form()
            st.rerun()


# =============================================================================
# Tab 1 - enter positions
# =============================================================================
def batch_top_level_form():
    """Several top-level (Level 1) posts at once: 3 rows to start, 'add more' for more."""
    n = st.session_state.l1_rows
    nz = st.session_state.nonce
    with st.form(f"l1_form_{nz}"):
        head = st.columns([3, 3, 2])
        head[0].markdown("**Position title**")
        head[1].markdown("**Person's name** (blank = vacant)")
        head[2].markdown("**Department**")
        for i in range(n):
            c = st.columns([3, 3, 2])
            c[0].text_input("Position title", key=f"l1_t_{nz}_{i}", label_visibility="collapsed",
                            placeholder="e.g. Director")
            c[1].text_input("Name", key=f"l1_n_{nz}_{i}", label_visibility="collapsed")
            c[2].text_input("Department", key=f"l1_d_{nz}_{i}", label_visibility="collapsed")
        saved = st.form_submit_button("Save top-level positions", type="primary")
    st.button("＋ Add 3 more rows", on_click=more_rows)

    if saved:
        pos, extra = st.session_state.pos, st.session_state.extra
        count = 0
        for i in range(n):
            title = st.session_state.get(f"l1_t_{nz}_{i}", "").strip()
            if not title:
                continue
            pos, extra, _ = du.add_position(
                pos, extra, title=title, name=st.session_state.get(f"l1_n_{nz}_{i}", ""),
                department=st.session_state.get(f"l1_d_{nz}_{i}", ""), reports_to="")
            count += 1
        if count:
            set_data(pos, extra)
            st.session_state.l1_rows = 3
            st.session_state.nonce += 1   # new widget keys = empty form
            st.session_state.flash = f"Added {count} top-level position(s). Now click ➕ beside a position to add posts below it."
            st.rerun()
        else:
            st.warning("Enter at least one position title.")


def position_form(mode: str, pid: str | None, parent: str):
    """Form to add a post under `parent`, or to edit post `pid`."""
    ss = st.session_state
    pos, extra = ss.pos, ss.extra
    by_id = du.positions_by_id(pos)
    labels = label_map(pos)
    cur = by_id.get(pid, {}) if mode == "edit" else {}

    if mode == "add":
        if parent:
            st.subheader("Add a position")
            st.caption(f"Reports to: **{labels.get(parent, parent)}**")
        else:
            st.subheader("Add a top-level position")
    else:
        st.subheader(f"Edit {pid}")

    departments = sorted({d for d in pos["department"] if d})
    blocked = set(du.descendants(pos, pid, include_root=True)) if mode == "edit" else set()
    extra_options = [i for i in by_id if i not in blocked]
    current_extra = [r["also_reports_to"] for r in extra.to_dict("records")
                     if r["position_id"] == pid and r["also_reports_to"] in extra_options] if mode == "edit" else []

    with st.form(f"pf_{mode}_{pid}_{parent}_{ss.nonce}"):
        title = st.text_input("Position title *", value=cur.get("position_title", ""))
        name = st.text_input("Person's name (leave blank if vacant)", value=cur.get("name", ""))
        dep_choice = st.selectbox(
            "Department", [""] + departments,
            index=([""] + departments).index(cur["department"]) if cur.get("department") in departments else 0)
        dep_new = st.text_input("…or type a new department", value="")
        vacant = st.checkbox("This post is vacant", value=(cur.get("status") == "Vacant"))
        c1, c2 = st.columns(2)
        email = c1.text_input("Email", value=cur.get("email", ""))
        phone = c2.text_input("Phone", value=cur.get("phone", ""))

        new_parent = parent
        sort_value = None
        if mode == "edit":
            parent_options = [""] + [i for i in by_id if i not in blocked]
            new_parent = st.selectbox(
                "Reports to (main / solid line)", parent_options,
                index=parent_options.index(cur["reports_to"]) if cur.get("reports_to") in parent_options else 0,
                format_func=lambda i: "(top level - reports to no one)" if i == "" else labels[i])
            order_now = cur.get("sort_order")
            sort_value = st.number_input(
                "Order among posts under the same manager (1 = first / leftmost)", min_value=0, step=1,
                value=int(order_now) if pd.notna(order_now) else 1)

        also = st.multiselect(
            "Also reports to (optional - dotted / second line)", extra_options, default=current_extra,
            format_func=lambda i: labels[i])
        rel = st.selectbox("Type of extra line (optional)", [""] + du.RELATIONSHIP_OPTIONS)

        b = st.columns(3)
        save = b[0].form_submit_button("Save", type="primary")
        save_more = b[1].form_submit_button("Save & add another") if mode == "add" else False
        cancel = b[2].form_submit_button("Cancel")

    if cancel:
        close_form()
        st.rerun()
    if save or save_more:
        if not title.strip():
            st.error("Position title is required.")
            return
        department = dep_new.strip() or dep_choice
        status = "Vacant" if (vacant or not name.strip()) else "Filled"
        also = [a for a in also if a != new_parent]  # main manager can't also be an extra manager
        if mode == "add":
            pos2, extra2, new_id = du.add_position(
                pos, extra, title=title, name=name, department=department, reports_to=parent,
                status=status, email=email, phone=phone, also_reports_to=also, relationship=rel)
            set_data(pos2, extra2)
            ss.flash = f"Added {new_id}: {title.strip()}"
            if not save_more:
                close_form()
        else:
            pos2, extra2 = du.update_position(
                pos, extra, pid, position_title=title, name=name, department=department,
                reports_to=new_parent, status=status, email=email, phone=phone,
                sort_order=sort_value, also_reports_to=also, relationship=rel)
            set_data(pos2, extra2)
            ss.flash = f"Updated {pid}"
            close_form()
        ss.nonce += 1
        st.rerun()


def delete_panel(pid: str):
    pos, extra = st.session_state.pos, st.session_state.extra
    labels = label_map(pos)
    below = du.descendants(pos, pid)
    st.subheader("Delete position")
    st.warning(f"Delete **{labels.get(pid, pid)}**?")
    with_subtree = False
    if below:
        choice = st.radio(
            f"{len(below)} position(s) work below this one:",
            ["Keep them - move them up to this post's manager", "Delete them as well"])
        with_subtree = choice == "Delete them as well"
    c1, c2 = st.columns(2)
    if c1.button("Yes, delete", type="primary", key="del_yes"):
        pos2, extra2 = du.delete_position(pos, extra, pid, with_subtree=with_subtree)
        set_data(pos2, extra2)
        st.session_state.flash = f"Deleted {pid}"
        close_form()
        st.rerun()
    if c2.button("Cancel", key="del_no"):
        close_form()
        st.rerun()


def tab_enter():
    ss = st.session_state

    if ss.pos.empty:
        st.subheader("Step 1 - Level 1 (top) positions")
        st.write("Enter the top-level positions of the institution. "
                 "Leave unused rows empty. Need more rows? Use **Add 3 more rows**. "
                 "You can also load an Excel file from the sidebar.")
        batch_top_level_form()
        return

    pos = ss.pos
    by_id = du.positions_by_id(pos)
    left, right = st.columns([3, 2])

    with left:
        top = st.columns([2, 2, 2])
        top[0].button("➕ Add top-level position", on_click=open_form, args=("add",))
        query = top[1].text_input("Search", placeholder="search title / name / department",
                                  label_visibility="collapsed")
        levels = du.compute_levels(pos)
        max_level = max(levels.values()) if levels else 1
        show_to = top[2].slider("Show levels up to", 1, max(2, max_level), max(2, max_level))

        with st.expander("Add several top-level positions at once"):
            batch_top_level_form()

        extra_children = {r["position_id"] for r in ss.extra.to_dict("records")}
        st.caption("➕ add a position below · ✏️ edit · 🗑️ delete")
        shown = 0
        q = query.strip().lower()
        for pid, depth in du.tree_rows(pos):
            r = by_id[pid]
            if depth + 1 > show_to:
                continue
            if q and q not in " ".join([r["position_title"], r["name"], r["department"], pid]).lower():
                continue
            shown += 1
            c = st.columns([8, 1, 1, 1], vertical_alignment="center")
            indent = "\u2003\u2003" * depth + ("└ " if depth else "")
            text = f"{indent}**{md_escape(r['position_title'])}**"
            if r["name"]:
                text += f" · {md_escape(r['name'])}"
            else:
                text += " · :red[vacant]"
            if r["department"]:
                text += f" · *{md_escape(r['department'])}*"
            if pid in extra_children:
                text += " 🔗"
            c[0].markdown(text + f"  `{pid}`")
            c[1].button("➕", key=f"add_{pid}", help="Add a position below this one",
                        on_click=open_form, args=("add",), kwargs={"parent": pid})
            c[2].button("✏️", key=f"edit_{pid}", help="Edit this position",
                        on_click=open_form, args=("edit",), kwargs={"pid": pid})
            c[3].button("🗑️", key=f"del_{pid}", help="Delete this position",
                        on_click=open_form, args=("delete",), kwargs={"pid": pid})
        if shown == 0:
            st.info("No positions match.")
        st.caption(f"{len(pos)} positions in total · 🔗 = has an additional reporting line")

    with right:
        with st.container(border=True):
            form = ss.form
            if form is None:
                st.markdown("#### How to add positions")
                st.markdown(
                    "1. Click **➕** beside a position to add a post *below* it.\n"
                    "2. Fill in the details and press **Save & add another** to add several "
                    "posts under the same manager quickly.\n"
                    "3. Use **✏️** to change details, move a post to another manager, or add a "
                    "second (dotted-line) manager.\n"
                    "4. Go to the **Chart & export** tab when you are ready."
                )
            elif form["mode"] == "delete" and form["pid"] in by_id:
                delete_panel(form["pid"])
            elif form["mode"] == "edit" and form["pid"] in by_id:
                position_form("edit", form["pid"], "")
            elif form["mode"] == "add" and (form["parent"] == "" or form["parent"] in by_id):
                position_form("add", None, form["parent"])
            else:
                close_form()
                st.rerun()


# =============================================================================
# Tab 2 - table / bulk edit
# =============================================================================
def _for_editor(df: pd.DataFrame) -> pd.DataFrame:
    """Show blanks as empty cells; keep number columns numeric."""
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].where(out[c] != "", None)
    return out


def tab_table():
    ss = st.session_state
    st.write("Edit like a spreadsheet, or paste rows copied from Excel. "
             "Click **Apply changes** below each table when finished. "
             "Leave `position_id` empty for new rows and it will be generated.")

    st.markdown("#### Positions")
    edited = st.data_editor(
        _for_editor(ss.pos), num_rows="dynamic", hide_index=True, key=f"pos_editor_{len(ss.pos)}_{ss.nonce}",
        column_config={
            "status": st.column_config.SelectboxColumn("status", options=du.STATUS_OPTIONS),
            "sort_order": st.column_config.NumberColumn("sort_order", min_value=0, step=1),
            "reports_to": st.column_config.TextColumn(
                "reports_to", help="position_id of the main manager. Blank = top level."),
        })
    if st.button("Apply changes to positions", type="primary"):
        new = du.clean_positions(edited)
        blank = new["position_id"] == ""
        for idx in new.index[blank]:
            new.loc[idx, "position_id"] = du.next_id(new)
        set_data(new, ss.extra)
        ss.nonce += 1
        errs, _ = du.validate(new, ss.extra)
        ss.flash = "Positions updated." + (f" {len(errs)} problem(s) found - see the Checks tab." if errs else "")
        st.rerun()

    st.markdown("#### Extra reporting lines (dotted / second manager)")
    ex_edit = st.data_editor(
        _for_editor(ss.extra), num_rows="dynamic", hide_index=True,
        key=f"extra_editor_{len(ss.extra)}_{ss.nonce}",
        column_config={
            "relationship": st.column_config.SelectboxColumn("relationship", options=du.RELATIONSHIP_OPTIONS),
        })
    if st.button("Apply changes to extra lines"):
        set_data(ss.pos, du.clean_extra(ex_edit))
        ss.nonce += 1
        ss.flash = "Extra reporting lines updated."
        st.rerun()


# =============================================================================
# Tab 3 - chart and export
# =============================================================================
def tab_chart():
    ss = st.session_state
    pos, extra = ss.pos, ss.extra
    errors, _ = du.validate(pos, extra)
    if pos.empty:
        st.info("Enter some positions first.")
        return
    if errors:
        st.error("Please fix these problems first (see the Checks tab for the full list):")
        for e in errors[:8]:
            st.write("• " + e)
        return

    levels = du.compute_levels(pos)
    max_level = max(levels.values())
    labels = label_map(pos)

    with st.expander("Chart settings", expanded=True):
        c1, c2, c3 = st.columns(3)
        mode_label = c1.radio(
            "What to show",
            ["Whole institution on one page", "One branch only", "Overview + one page per branch (best for 50+ posts)"])
        mode = {"Whole": "full", "One b": "branch", "Overv": "split"}[mode_label[:5]]

        root_id, split_level = None, 2
        if mode == "branch":
            root_id = c1.selectbox("Show this post and everything below", list(labels),
                                   format_func=lambda i: f"L{levels[i]}  {labels[i]}")
        if mode == "split":
            if max_level < 2:
                c1.info("Only one level exists, so a single page is used.")
            split_level = c1.number_input("Overview shows levels 1 to", 1, max(1, max_level - 1),
                                          min(2, max(1, max_level - 1)))

        limit = c2.checkbox("Limit the number of levels drawn", value=False)
        max_depth = c2.slider("Levels drawn", 1, max(2, max_level), min(3, max(2, max_level))) if limit else None
        direction = c2.radio("Direction", ["Top to bottom", "Left to right"], horizontal=True)
        line = c2.selectbox("Line style", ["Straight", "Curved", "Elbow (experimental)"])

        color = c3.checkbox("Colour by department", value=True)
        vac = c3.checkbox("Highlight vacant posts", value=True)
        show_id = c3.checkbox("Show position IDs", value=False)
        contact = c3.checkbox("Show email / phone in boxes", value=False)
        legend = c3.checkbox("Show legend", value=True)

    opts = ChartOptions(
        rankdir="TB" if direction == "Top to bottom" else "LR",
        line_style={"Straight": "polyline", "Curved": "spline", "Elbow (experimental)": "ortho"}[line],
        color_by_department=color, highlight_vacant=vac, show_id=show_id,
        show_contact=contact, show_legend=legend, institution=ss.get("institution", "").strip(),
    )
    pages = make_pages(pos, extra, mode=mode, root_id=root_id, max_depth=max_depth,
                       split_level=int(split_level), opts=opts)
    if not pages:
        st.info("Nothing to draw.")
        return

    m = st.columns(3)
    m[0].metric("Positions", len(pos))
    m[1].metric("Levels", max_level)
    m[2].metric("Chart pages", len(pages))

    biggest = max(p["count"] for p in pages)
    if biggest > 40:
        st.warning(f"One page has {biggest} boxes, so text will be small when printed. "
                   "Try 'Overview + one page per branch', 'Left to right', or A2 paper.")

    idx = 0
    if len(pages) > 1:
        idx = st.selectbox("Preview page", range(len(pages)),
                           format_func=lambda i: f"{i + 1}. {pages[i]['title']} ({pages[i]['count']} boxes)")
    st.graphviz_chart(pages[idx]["graph"].source)

    # ---- export ---------------------------------------------------------------
    st.divider()
    st.subheader("Print / download")
    e1, e2 = st.columns(2)
    paper = e1.selectbox("Paper size", list(PAPERS), index=1,
                         help="Chart pages are fitted to this size. Tall charts use portrait automatically.")
    with_table = e2.checkbox("Include the position details table in the Word file", value=True)

    sig = hashlib.md5(repr((pos.to_csv(), extra.to_csv(), mode, root_id, max_depth, split_level,
                            opts, paper, with_table)).encode("utf-8")).hexdigest()
    if st.button("Prepare PDF and Word files", type="primary"):
        try:
            with st.spinner("Creating files…"):
                inst = opts.institution
                ss.exports = {
                    "sig": sig,
                    "pdf": pages_to_pdf(pages, paper),
                    "docx": build_word(pages, pos, extra, paper, inst, include_table=with_table),
                    "png": graph_to_png(pages[idx]["graph"]),
                    "name": f"{slug(inst)}_{datetime.date.today().isoformat()}",
                    "png_page": idx + 1,
                }
        except Exception as exc:
            ss.exports = None
            st.error(
                f"Could not create the files: {exc}\n\n"
                "If the message mentions 'dot' or 'ExecutableNotFound', the Graphviz program is "
                "missing - add a file called packages.txt containing the word graphviz to your repository.")

    ex = ss.exports
    if ex and ex["sig"] == sig:
        d = st.columns(3)
        d[0].download_button("⬇ PDF", ex["pdf"], file_name=f"{ex['name']}.pdf", mime="application/pdf")
        d[1].download_button(
            "⬇ Word (.docx)", ex["docx"], file_name=f"{ex['name']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        d[2].download_button(f"⬇ Picture of page {ex['png_page']} (PNG)", ex["png"],
                             file_name=f"{ex['name']}_page{ex['png_page']}.png", mime="image/png")
    elif ex:
        st.info("Data or settings changed - press **Prepare PDF and Word files** again.")


# =============================================================================
# Tab 4 - checks
# =============================================================================
def tab_checks():
    ss = st.session_state
    errors, warns = du.validate(ss.pos, ss.extra)
    if ss.pos.empty:
        st.info("No data yet.")
        return
    if not errors:
        st.success("No errors found. The data is ready for charting.")
    for e in errors:
        st.error(e)
    for w in warns:
        st.warning(w)

    if not errors:
        levels = du.compute_levels(ss.pos)
        by_level = pd.Series(levels).value_counts().sort_index()
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Positions per level**")
            st.bar_chart(by_level.rename_axis("Level").rename("Positions"))
        with c2:
            st.markdown("**Summary**")
            vacant = int(((ss.pos["status"] == "Vacant") | (ss.pos["name"] == "")).sum())
            st.write(f"- Total positions: **{len(ss.pos)}**")
            st.write(f"- Vacant: **{vacant}**")
            st.write(f"- Extra reporting lines: **{len(ss.extra)}**")
            dept = ss.pos[ss.pos["department"] != ""]["department"].value_counts()
            if not dept.empty:
                st.markdown("**Positions per department**")
                st.dataframe(dept.rename("Positions"))


# =============================================================================
# Main
# =============================================================================
init_state()
sidebar()
if st.session_state.get("flash"):
    st.toast(st.session_state.pop("flash"), icon="✅")
st.title("Institution Organogram Builder")
tabs = st.tabs(["1. Enter positions", "2. Table view / bulk edit", "3. Chart & export", "4. Checks"])
with tabs[0]:
    tab_enter()
with tabs[1]:
    tab_table()
with tabs[2]:
    tab_chart()
with tabs[3]:
    tab_checks()
