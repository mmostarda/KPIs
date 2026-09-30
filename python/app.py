"""Interfaccia grafica di kpimeta.

Avvio (dalla cartella python/):   streamlit run app.py
Flusso: 1 DB -> 2 Sorgente -> 3 Metadati -> 4 Revisione -> 5 Scrittura
"""

from __future__ import annotations

import datetime as dt
import os
import sys
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="KPI metadata", page_icon="📊", layout="wide")
# the first start loads pandas & co. (tens of seconds on a PC with an antivirus): say so
# instead of showing an empty page; the imports below come after this on purpose
_loading = None if "kpimeta.excel_db" in sys.modules else st.empty()
if _loading is not None:
    _loading.info("Avvio in corso: caricamento delle librerie (al primo avvio può richiedere un minuto)…")

import pandas as pd

from kpimeta.batch import (BatchError, assign_ids, batch_from_excel, batch_from_frame, batch_manual, compose,
                           read_excel_table)
from kpimeta.config import ConfigError, backup_root, load_prefs, load_settings, save_prefs
from kpimeta.excel_db import (DBError, add_support_value, date_report, overwrite_metadata, preview_write, read_db,
                              write_batch)
from kpimeta.issues import ERROR, INFO, WARNING
from kpimeta.lola import LolaError, export_all_db, is_lola_output, read_lola_files
from kpimeta.mapping import MappingError, load_mapping
from kpimeta.schema import load_schema
from kpimeta.ui_helpers import (FILE_COL, FROM_COL, MIXED_COL, PROBLEMS_COL, ROWS_COL, SOURCE_COL,
                                apply_edits, apply_group_edits, batch_excel_bytes, diff_frame, display_frame,
                                field_value, form_defaults, inferred_columns, inferred_frame, inferred_groups,
                                issues_frame, missing_support_values, one_row_excel_bytes, problems_text,
                                values_to_widgets, widget_key)
from kpimeta.validate import (KpiDecision, canonicalize_choices, check_decisions, duplicates_vs_db,
                              internal_duplicates, kpi_resolution, validate_batch)
from kpimeta.values import is_missing, to_text

STEPS = ["1 · DB", "2 · Sorgente", "3 · Metadati", "4 · Revisione", "5 · Scrittura"]
SOURCES = {
    "lola": "Output LoLa",
    "excel": "Excel Metadata + KPI",
    "manual": "Nuovo test (1 riga)",
    "edit": "Modifica un ID esistente",
}
CHOOSE, ADD, DISCARD, MAP = "— scegli —", "Aggiungi come nuova colonna", "Scarta", "→ "
DATE_MIN, DATE_MAX = dt.date(1900, 1, 1), dt.date(2100, 12, 31)

if _loading is not None:
    _loading.empty()
ss = st.session_state


@st.cache_resource
def get_config():
    settings = load_settings()
    return settings, load_schema(settings.schema_path)


try:
    settings, schema = get_config()
except ConfigError as exc:
    st.error(f"Configurazione non valida: {exc}")
    st.stop()


# --- state -----------------------------------------------------------------------

def init_state() -> None:
    prefs = load_prefs()
    defaults = {
        "step": STEPS[0], "db_path": prefs.get("last_db", ""), "snapshot": None, "source": "lola",
        "lola_dir": prefs.get("last_lola_dir", ""), "excel_path": prefs.get("last_excel", ""),
        "mapping_path": str(settings.lola_mapping_path), "base": None, "source_report": None,
        "batch": None, "view_rows": None, "editor_version": 0, "edit_record": None, "edit_values": None,
        "result": None, "form_ready": False, "form_mode": "overwrite", "infer_filename": True,
        "infer_template": False, "manual_kpis": [], "manual_new_kpis": "", "flash": [],
        "clear_empty_fields": False, "template_choice": "", "form_import_path": "", "edit_id_input": 1,
        "export_all_db": False, "problem_cols_only": False,
    }
    for key, value in defaults.items():
        ss.setdefault(key, value)


PERSISTENT_KEYS = ("db_path", "source", "lola_dir", "excel_path", "mapping_path", "form_mode", "infer_filename",
                   "infer_template", "manual_kpis", "manual_new_kpis", "clear_empty_fields", "template_choice",
                   "form_import_path", "edit_id_input", "export_all_db", "review_cols")
PERSISTENT_PREFIXES = ("f::", "kpi::", "lola_files::")


def keep_widget_values() -> None:
    """Streamlit forgets the value of widgets that are not drawn in the current step:
    re-assigning them keeps the form filled when moving between steps."""
    for key in list(ss.keys()):
        if key in PERSISTENT_KEYS or str(key).startswith(PERSISTENT_PREFIXES):
            ss[key] = ss[key]


def flash(kind: str, message: str) -> None:
    ss.flash.append((kind, message))


def show_flash() -> None:
    for kind, message in ss.flash:
        getattr(st, kind)(message)
    ss.flash = []


def goto(index: int) -> None:
    ss.step = STEPS[index]


def snapshot():
    return ss.snapshot


def meta_headers() -> list[str]:
    snap = snapshot()
    return snap.metadata.headers if snap is not None and snap.metadata is not None else []


def options() -> dict[str, list[str]]:
    snap = snapshot()
    return snap.options(schema) if snap is not None else {f.column: list(f.choices) for f in schema.fields}


def connect(path: str) -> None:
    path = path.strip().strip('"')
    if not path:
        flash("error", "Indica il percorso del DB Excel.")
        return
    try:
        with st.spinner("Lettura del DB…"):
            ss.snapshot = read_db(path, schema, settings)
        save_prefs(last_db=path)
    except (DBError, OSError, ValueError, KeyError) as exc:
        ss.snapshot = None
        flash("error", f"Impossibile leggere il DB: {exc}")


def refresh_if_changed() -> None:
    snap = snapshot()
    if snap is None:
        return
    try:
        stat = snap.path.stat()
    except OSError:
        return
    if (stat.st_mtime_ns, stat.st_size) != (snap.mtime_ns, snap.size):
        flash("info", "Il DB è stato modificato: riletto.")
        connect(str(snap.path))


# --- file dialogs (only when the app runs on the user's PC) -------------------------

def can_browse() -> bool:
    return os.name == "nt" or bool(os.environ.get("DISPLAY"))


def browse(kind: str, initial: str = "") -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        start = str(Path(initial).parent if initial and Path(initial).is_file() else (initial or Path.home()))
        if kind == "folder":
            return filedialog.askdirectory(parent=root, initialdir=start) or None
        return filedialog.askopenfilename(parent=root, initialdir=start,
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tutti", "*.*")]) or None
    finally:
        root.destroy()


def browse_into(key: str, kind: str) -> None:
    chosen = browse(kind, ss.get(key, ""))
    if chosen:
        ss[key] = chosen


def path_input(label: str, key: str, kind: str = "file", placeholder: str = "") -> str:
    cols = st.columns([8, 1], vertical_alignment="bottom") if can_browse() else [st.container()]
    cols[0].text_input(label, key=key, placeholder=placeholder)
    if can_browse():
        cols[1].button("Sfoglia…", key=f"browse::{key}", on_click=browse_into, args=(key, kind))
    return ss[key].strip().strip('"')


def nav(back: int | None = None, forward: int | None = None, label: str = "Avanti →",
        disabled: bool = False, on_forward=None, on_back=None) -> None:
    st.divider()
    row = st.container(horizontal=True)
    if back is not None:
        row.button("← Indietro", key=f"back::{ss.step}", on_click=on_back or goto,
                   args=() if on_back else (back,))
    if forward is not None:
        row.button(label, key=f"next::{ss.step}", type="primary", disabled=disabled,
                   on_click=on_forward or goto, args=() if on_forward else (forward,))


def require_db() -> bool:
    if snapshot() is None:
        st.info("Collega prima il DB Excel al passo 1.")
        st.button("Vai al passo 1", on_click=goto, args=(0,))
        return False
    return True


# --- step 1: DB --------------------------------------------------------------------

def step_db() -> None:
    st.header("1 · DB Excel")
    st.caption("Il file letto da Power BI: nome, fogli, tabelle (TableMet, TableKPI) e colonne restano invariati.")
    path = path_input("Percorso del DB", "db_path", placeholder=r"C:\...\KPI_DB.xlsx")
    if st.button("Connetti / rileggi", type="primary"):
        connect(path)
        st.rerun()  # refresh the sidebar status too
    show_flash()
    snap = snapshot()
    if snap is None:
        with st.expander("Come funziona", expanded=True):
            st.markdown(
                "1. **DB**: colleghi il file Excel del DB (viene solo letto).\n"
                "2. **Sorgente**: output LoLa, un Excel Metadata + KPI, un nuovo test o la modifica di un ID.\n"
                "3. **Metadati**: form con template, liste del DB e deduzione dal nome file.\n"
                "4. **Revisione**: tabella modificabile con i problemi evidenziati riga per riga.\n"
                "5. **Scrittura**: controlli finali, backup automatico e una sola scrittura atomica.")
        return

    c = st.columns(5)
    c[0].metric("Righe Metadata", len(snap.metadata.df) if snap.metadata is not None else "—")
    c[1].metric("Righe KPIs", len(snap.kpis.df) if snap.kpis is not None else "—")
    c[2].metric("Prossimo ID", snap.next_id)
    c[3].metric("Colonne KPI", len(snap.kpi_names))
    c[4].metric("Template", 0 if snap.templates is None else len(snap.templates))
    blockers = [i for i in snap.issues if i.severity == ERROR]
    for issue in blockers:
        st.error(issue.message)
    for issue in (i for i in snap.issues if i.severity == WARNING):
        st.warning(issue.message)
    infos = [i for i in snap.issues if i.severity == INFO]
    if not blockers:
        st.success(f"DB coerente ({snap.path.name}): si può scrivere."
                   + (" Chiudi però il file in Excel prima della scrittura." if snap.lock else ""))
    if infos:
        with st.expander(f"Note ({len(infos)})"):
            for issue in infos:
                st.write("• " + issue.message)
    with st.expander("Come sono salvate oggi le colonne data"):
        st.dataframe(date_report(snap, schema), hide_index=True)
    with st.expander("Visualizza il DB"):
        tabs = st.tabs(["Metadata", "KPIs"])
        for tab, info in zip(tabs, (snap.metadata, snap.kpis)):
            if info is not None:
                tab.dataframe(info.df.map(lambda v: to_text(v) if isinstance(v, (dt.date, dt.datetime)) else v),
                              hide_index=True)
    nav(forward=1)


# --- step 2: source ------------------------------------------------------------------

def reset_batch() -> None:
    ss.batch = None
    ss.view_rows = None
    ss.review_cols = None
    ss.problem_cols_only = False
    ss.editor_version += 1
    ss.result = None


def step_source() -> None:
    st.header("2 · Sorgente delle righe")
    if not require_db():
        return
    st.radio("Da dove arrivano le righe?", list(SOURCES), format_func=SOURCES.get, key="source", horizontal=True)
    {"lola": source_lola, "excel": source_excel, "manual": source_manual, "edit": source_edit}[ss.source]()


def _lola_candidates(folder: Path) -> list[Path]:
    return sorted(p for p in folder.glob("*.xls*") if is_lola_output(p))


def source_lola() -> None:
    mapping_path = path_input("Mapping LoLa → DB (CSV)", "mapping_path")
    try:
        mapping = load_mapping(mapping_path)
        st.caption(f"{len(mapping.entries)} colonne LoLa mappate ({Path(mapping_path).name})")
    except MappingError as exc:
        st.error(str(exc))
        return
    folder_text = path_input("Cartella con gli output LoLa", "lola_dir", kind="folder")
    folder = Path(folder_text) if folder_text else None
    files = _lola_candidates(folder) if folder and folder.is_dir() else []
    if folder_text and not files:
        st.warning("Nessun file .xlsx trovato nella cartella.")
    names = [p.name for p in files]
    files_key = f"lola_files::{folder_text}"
    ss[files_key] = [n for n in ss.get(files_key, names) if n in names]
    selected = st.multiselect("File da leggere", names, key=files_key)
    export = st.checkbox("Crea anche ALL_DB.xlsx nella cartella (file intermedio dell'app MATLAB)",
                         key="export_all_db")
    if st.button("Leggi i file LoLa", type="primary", disabled=not selected):
        try:
            with st.spinner("Lettura degli output LoLa…"):
                frame, reports = read_lola_files([folder / name for name in selected], mapping, settings)
                ss.base = batch_from_frame(frame, schema, settings, "lola", meta_headers())
                ss.source_report = pd.DataFrame([{
                    "File": Path(r.path).name, "Righe": r.data_rows,
                    "Scartate (senza File Name)": r.dropped_without_filename,
                    "Colonne del mapping assenti": ", ".join(r.lola_names_missing),
                    "Colonne ignorate": ", ".join(r.unmapped_headers),
                    "Avvisi": "; ".join(r.warnings)} for r in reports])
                if export:
                    export_all_db(frame, folder / "ALL_DB.xlsx")
                    flash("info", f"Creato {folder / 'ALL_DB.xlsx'}")
            ss.form_mode = "overwrite"
            reset_batch()
            save_prefs(last_lola_dir=folder_text)
        except (LolaError, OSError, ValueError) as exc:
            ss.base = None  # never keep the rows of a previous read after a failed one
            reset_batch()
            flash("error", f"Lettura non riuscita: {exc}")
    show_flash()
    base = ss.base
    if base is not None and base.source == "lola":
        st.success(f"{base.n_rows} righe lette, {len(base.kpi_columns)} colonne KPI: {', '.join(base.kpi_columns)}")
        if ss.source_report is not None:
            st.dataframe(ss.source_report, hide_index=True)
    nav(back=0, forward=2, disabled=base is None or base.source != "lola")


def source_excel() -> None:
    st.caption("Un file con una riga per test: colonne dei metadati (nomi del DB) + colonne KPI. "
               "Esempio: il Metadata_ALL_DB.xlsx creato dall'app MATLAB o esportato al passo 4.")
    path = path_input("File Excel", "excel_path")
    if st.button("Leggi il file", type="primary", disabled=not path):
        try:
            with st.spinner("Lettura…"):
                ss.base = batch_from_excel(path, schema, settings, meta_headers())
            ss.form_mode = "fill"
            reset_batch()
            save_prefs(last_excel=path)
        except (BatchError, OSError, ValueError, KeyError) as exc:
            ss.base = None
            reset_batch()
            flash("error", f"Lettura non riuscita: {exc}")
    show_flash()
    base = ss.base
    if base is not None and base.source == "excel":
        filled = [c for c in base.meta_columns if base.df[c].map(lambda v: not is_missing(v)).any()]
        st.success(f"{base.n_rows} righe: {len(filled)} colonne di metadati con valori, "
                   f"{len(base.kpi_columns)} colonne KPI")
        for note in base.notes:
            st.caption("• " + note.message)
    nav(back=0, forward=2, disabled=base is None or base.source != "excel")


def start_manual() -> None:
    extra = [k.strip() for k in ss.get("manual_new_kpis", "").split(";") if k.strip()]
    ss.base = batch_manual(schema, {}, list(ss.manual_kpis) + extra)
    reset_batch()
    goto(2)


def source_manual() -> None:
    snap = snapshot()
    st.caption("Una riga: i metadati si compilano al passo 3, i valori dei KPI nella tabella del passo 4.")
    st.multiselect("KPI da inserire", snap.kpi_names, key="manual_kpis")
    st.text_input("KPI nuovi, non ancora nel DB (separati da ;)", key="manual_new_kpis")
    nav(back=0, forward=2, on_forward=start_manual)


def load_edit_record() -> None:
    record = snapshot().metadata_row(ss.edit_id_input, settings.id_column)
    if record is None:
        flash("error", f"ID {ss.edit_id_input} non trovato in {settings.metadata_table}.")
        ss.edit_record = None
        return
    ss.edit_record = record
    ss.update(form_defaults(schema, with_defaults=False))
    ss.update(values_to_widgets(schema, record))
    ss.form_ready = True
    ss.base = None
    reset_batch()


def source_edit() -> None:
    row = st.columns([2, 2, 6], vertical_alignment="bottom")
    row[0].number_input("ID da modificare", min_value=1, step=1, key="edit_id_input")
    row[1].button("Carica nel form", on_click=load_edit_record)
    show_flash()
    record = ss.edit_record
    if record is not None:
        st.success(f"ID {to_text(record.get(settings.id_column))}: {to_text(record.get('IDName')) or ''} · "
                   f"{to_text(record.get('Date')) or ''} · {to_text(record.get('FileName')) or ''}")
    nav(back=0, forward=2, disabled=record is None)


# --- step 3: common metadata form -------------------------------------------------------

def template_column():
    templates = snapshot().templates
    if templates is None:
        return None
    return next((c for c in templates.columns if str(c).lower() == settings.template_name_column.lower()), None)


def template_names() -> list[str]:
    column = template_column()
    return [] if column is None else [n for n in (to_text(v) for v in snapshot().templates[column]) if n]


def form_values() -> dict:
    values = {}
    for f in schema.fields:
        if f.type == "id":
            continue
        value = field_value(f, ss.get(widget_key(f.column)))
        if value is not None:
            values[f.column] = value
    return values


def apply_template(name: str) -> None:
    column = template_column()
    if column is None:
        flash("error", f"Colonna {settings.template_name_column} non trovata in {settings.templates_sheet}.")
        return
    templates = snapshot().templates
    rows = templates[templates[column].map(lambda v: (to_text(v) or "") == name)]
    if not rows.empty:
        ss.update(values_to_widgets(schema, rows.iloc[0].to_dict()))
        flash("success", f"Template '{name}' applicato al form.")


def import_form_row(path: str) -> None:
    try:
        frame, _ = read_excel_table(path)
    except (BatchError, OSError, ValueError) as exc:
        flash("error", f"Lettura non riuscita: {exc}")
        return
    if len(frame) != 1:
        flash("error", f"Il file deve contenere esattamente una riga (trovate {len(frame)}).")
        return
    ss.update(values_to_widgets(schema, frame.iloc[0].to_dict()))
    flash("success", "Valori importati nel form.")


def clear_form() -> None:
    ss.update(form_defaults(schema, with_defaults=False))


def compose_and_review() -> None:
    values = form_values()
    snap = snapshot()
    if ss.source == "edit":
        ss.edit_values = values
        goto(3)
        return
    base = ss.base
    if base.source == "manual":
        batch = batch_manual(schema, values, base.kpi_columns)
    else:
        batch = compose(base, schema, settings, values, mode=ss.form_mode, templates=snap.templates,
                        infer_template=ss.infer_template, infer_filename=ss.infer_filename)
    canonicalize_choices(batch, schema, options())
    assign_ids(batch, snap.next_id, settings.id_column)
    ss.batch = batch
    ss.view_rows = None
    ss.review_cols = None  # visible columns recomputed for the new batch
    ss.problem_cols_only = False
    ss.editor_version += 1
    goto(3)


def add_value_popover(field, label: str = "+", value: str = "", scope: str = "form") -> None:
    """Add a value to the support sheet of a choice field. In the form the new value is
    selected; in the review (`value` prefilled) the rows are checked again."""
    support = snapshot().support.get((field.choices_sheet, field.choices_column))
    if support is None or support.problem:
        return
    with st.popover(label, help=f"Aggiungi un valore al foglio {field.choices_sheet} del DB",
                    key=f"pop::{scope}::{field.column}"):
        values = {}
        for i, header in enumerate(support.headers):
            if i == 0 and header != support.header:
                st.caption(f"{header}: assegnato automaticamente")
                continue
            # default value, not session state: the content of a popover is created only when it opens
            values[header] = st.text_input(header, value=value if header == support.header else "",
                                           key=f"add::{scope}::{field.column}::{header}")
        if st.button("Aggiungi al DB", key=f"addbtn::{scope}::{field.column}", type="primary"):
            new_value = (values.get(support.header) or "").strip()
            if not new_value:
                st.error(f"Indica un valore per {support.header}")
                return
            try:
                with st.spinner("Scrittura nel foglio di supporto…"):
                    add_support_value(snapshot().path, field.choices_sheet,
                                      {k: v for k, v in values.items() if v.strip()}, settings)
                connect(str(snapshot().path))
                if scope == "form":
                    ss[f"pending::{widget_key(field.column)}"] = new_value
                else:
                    flash("success", f"'{new_value}' aggiunto al foglio {field.choices_sheet}")
                st.rerun()
            except DBError as exc:
                st.error(str(exc))


def render_field(field, locked: str | None) -> None:
    key = widget_key(field.column)
    help_text = field.notes or None
    if field.type == "id":
        edited = ss.edit_record.get(settings.id_column) if ss.source == "edit" and ss.edit_record else None
        st.text_input(field.label, value=to_text(edited) or f"automatico (da {snapshot().next_id})",
                      disabled=True)
        return
    if locked:
        help_text = locked
    disabled = bool(locked)
    if field.type == "text":
        st.text_input(field.label, key=key, help=help_text, disabled=disabled)
    elif field.type in ("int", "number"):
        integer = field.type == "int"
        cast = int if integer else float
        st.number_input(field.label, key=key, help=help_text, disabled=disabled,
                        min_value=cast(field.min) if field.min is not None else None,
                        max_value=cast(field.max) if field.max is not None else None,
                        step=1 if integer else None, format="%d" if integer else "%g")
    elif field.type == "date":
        st.date_input(field.label, key=key, help=help_text, disabled=disabled, format="DD/MM/YYYY",
                      min_value=DATE_MIN, max_value=DATE_MAX)
    elif field.type == "choice":
        known = list(field.choices) or options().get(field.column, [])
        choices = ["", *known]
        current = ss.get(key, "")
        if current and current not in choices:
            choices.append(current)
        inner = st.columns([5, 1], vertical_alignment="bottom") if field.uses_support_sheet else [st.container()]
        inner[0].selectbox(field.label, choices, key=key, help=help_text, disabled=disabled,
                           format_func=lambda v: v if v in known or not v else f"{v} (non in lista)")
        if field.uses_support_sheet:
            with inner[1]:
                add_value_popover(field)
    elif field.type == "multichoice":
        current = [v for v in ss.get(key, []) if v not in field.choices]
        st.multiselect(field.label, [*field.choices, *current], key=key, help=help_text, disabled=disabled)


def step_form() -> None:
    st.header("3 · Metadati")
    if not require_db():
        return
    if ss.source != "edit" and ss.base is None:
        st.info("Scegli prima la sorgente al passo 2.")
        nav(back=1)
        return
    if not ss.form_ready:
        ss.update(form_defaults(schema, with_defaults=True))
        ss.form_ready = True
    for key in [k for k in ss if str(k).startswith("pending::")]:
        ss[key.removeprefix("pending::")] = ss.pop(key)

    source = ss.source
    bar = st.columns([3, 2, 3, 2], vertical_alignment="bottom")
    chosen = bar[0].selectbox("Template (Metadata_Templates)", ["", *template_names()], key="template_choice")
    bar[1].button("Applica template", disabled=not chosen, on_click=apply_template, args=(chosen,),
                  width="stretch")
    bar[2].text_input("Importa i valori da un Excel di una riga", key="form_import_path")
    bar[3].button("Importa nel form", disabled=not ss.form_import_path, width="stretch",
                  on_click=import_form_row, args=(ss.form_import_path.strip().strip('"'),))
    actions = st.container(horizontal=True)
    actions.download_button("Scarica il form come Excel", one_row_excel_bytes(form_values()),
                            file_name="metadati_form.xlsx")
    actions.button("Svuota il form", on_click=clear_form)
    show_flash()

    if source in ("lola", "excel"):
        with st.container(border=True):
            st.radio("I valori del form", ["overwrite", "fill"], key="form_mode", horizontal=True,
                     format_func={"overwrite": "sovrascrivono quelli della sorgente",
                                  "fill": "riempiono solo le celle vuote"}.get)
            c = st.columns(2)
            c[0].checkbox("Deduci Date, Driving_Mode, ManeuvreName e Driver dal nome file", key="infer_filename",
                          help="Regola attuale: X_X_AAAAMMGG_X_DrivingMode_Maneuvre_Driver_X.mf4")
            c[1].checkbox("Applica per ogni riga il template dei primi 2 token del nome file", key="infer_template")
            st.caption("Precedenza: colonne protette della sorgente (Repetition_ID, FileName, StartSpeed se "
                       "presente) > nome file > template > form. I campi vuoti non sovrascrivono nulla.")

    locked_columns = {}
    if source in ("lola", "excel"):
        for f in schema.fields:
            if f.lola_policy == "keep":
                locked_columns[f.column] = "Preso dalla sorgente (non modificabile qui)"
    for tab, tab_name in zip(st.tabs([f"{t} Metadata" for t in schema.tabs()]), schema.tabs()):
        with tab:
            for section in schema.sections(tab_name):
                st.markdown(f"**{section}**")
                cols = st.columns(3)
                for i, field in enumerate(schema.fields_in(tab_name, section)):
                    with cols[i % 3]:
                        render_field(field, locked_columns.get(field.column))
    label = "Avanti: rivedi le modifiche →" if source == "edit" else "Avanti: componi le righe →"
    nav(back=1, forward=3, label=label, on_forward=compose_and_review)


# --- step 4: review ------------------------------------------------------------------------

def editor_key() -> str:
    return f"editor::{ss.editor_version}"


def inferred_key() -> str:
    return f"inferred::{ss.editor_version}"


def view_rows() -> list[int]:
    return ss.view_rows if ss.view_rows is not None else list(range(ss.batch.n_rows))


def current_batch():
    state = ss.get(editor_key()) or {}
    return apply_edits(ss.batch, schema, view_rows(), state.get("edited_rows", {}), settings.id_column)


def batch_issues(batch):
    snap = snapshot()
    issues = validate_batch(batch, schema, settings, options()) + internal_duplicates(batch, schema)
    if snap.metadata is not None:
        issues += duplicates_vs_db(batch, schema, snap.metadata.df)
    return issues


def commit_edits() -> None:
    if ss.batch is not None:
        ss.batch = current_batch()
    ss.editor_version += 1


def apply_inferred_edits() -> None:
    """An edit of the inferred-values table goes at once to every row of the file; the
    pending edits of the main table are committed with it (both tables are re-created)."""
    edited = (ss.get(inferred_key()) or {}).get("edited_rows") or {}
    current = current_batch()
    ss.batch = apply_group_edits(current, schema, inferred_groups(current), edited)
    ss.editor_version += 1


def toggle_problem_rows() -> None:
    commit_edits()
    if ss.view_rows is None:
        rows = sorted({i.row for i in batch_issues(ss.batch)
                       if i.row is not None and i.severity in (ERROR, WARNING)})
        ss.view_rows = rows or None
        if not rows:
            flash("info", "Nessuna riga con problemi.")
    else:
        ss.view_rows = None


def toggle_problem_columns() -> None:
    commit_edits()
    if ss.problem_cols_only:
        ss.review_cols = None  # back to the default columns
        ss.problem_cols_only = False
        return
    with_issue = {i.column for i in batch_issues(ss.batch) if i.column and i.severity in (ERROR, WARNING)}
    columns = [c for c in ss.batch.review_columns() if c in with_issue]
    if columns:
        ss.review_cols = columns
        ss.problem_cols_only = True
    else:
        flash("info", "Nessuna colonna con problemi.")


def accept_invalid_as_empty() -> None:
    commit_edits()
    ss.batch.conv_issues.clear()


def column_config(batch) -> dict:
    cfg = {
        "_index": st.column_config.NumberColumn("riga", disabled=True, format="%d"),
        PROBLEMS_COL: st.column_config.TextColumn(PROBLEMS_COL, disabled=True, width="large"),
        settings.id_column: st.column_config.NumberColumn("ID", disabled=True, format="%d",
                                                          help="ID provvisorio: quello definitivo è "
                                                               "calcolato al momento della scrittura"),
        SOURCE_COL: st.column_config.TextColumn("File LoLa", disabled=True),
    }
    opts = options()
    for column in batch.meta_columns:
        f = schema.get(column)
        if f is None:
            cfg[column] = st.column_config.TextColumn(column, help="colonna di TableMet fuori dallo schema")
            continue
        if f.type == "id":
            continue
        locked = batch.source == "lola" and f.lola_policy == "keep"
        help_text = f"{f.label}" + (f" · {f.notes}" if f.notes else "")
        if f.type == "int":
            cfg[column] = st.column_config.NumberColumn(column, help=help_text, disabled=locked, step=1,
                                                        format="%d", min_value=f.min, max_value=f.max)
        elif f.type == "number":
            cfg[column] = st.column_config.NumberColumn(column, help=help_text, disabled=locked,
                                                        min_value=f.min, max_value=f.max)
        elif f.type == "date":
            cfg[column] = st.column_config.DateColumn(column, help=help_text, disabled=locked,
                                                      format="DD/MM/YYYY", min_value=DATE_MIN, max_value=DATE_MAX)
        elif f.type == "choice":
            present = [v for v in batch.df[column].dropna().unique() if isinstance(v, str)]
            allowed = list(f.choices) or opts.get(column, [])
            cfg[column] = st.column_config.SelectboxColumn(
                column, help=help_text, disabled=locked, options=list(dict.fromkeys([*allowed, *present])))
        else:
            cfg[column] = st.column_config.TextColumn(
                column, disabled=locked,
                help=help_text + (" · valori separati da ;" if f.type == "multichoice" else ""))
    for column in batch.kpi_columns:
        cfg[column] = st.column_config.NumberColumn(column, help="KPI")
    return cfg


def default_columns(batch, issues) -> list[str]:
    with_issue = {i.column for i in issues if i.column}
    keep = []
    for column in batch.review_columns():
        f = schema.get(column)
        if column in batch.kpi_columns or column in with_issue or (f is not None and f.type == "id"):
            keep.append(column)
        elif batch.df[column].map(lambda v: not is_missing(v)).any():
            keep.append(column)
    return keep


def step_review() -> None:
    st.header("4 · Revisione")
    if not require_db():
        return
    if ss.source == "edit":
        review_edit()
        return
    if ss.batch is None:
        st.info("Componi prima le righe al passo 3.")
        nav(back=2)
        return
    current = current_batch()
    issues = batch_issues(current)
    n_err = sum(1 for i in issues if i.severity == ERROR)
    n_warn = sum(1 for i in issues if i.severity == WARNING)
    c = st.columns(4)
    c[0].metric("Righe", current.n_rows)
    c[1].metric("Errori", n_err)
    c[2].metric("Avvisi", n_warn)
    c[3].metric("Colonne KPI", len(current.kpi_columns))

    actions = st.container(horizontal=True)
    actions.button("Mostra tutte le righe" if ss.view_rows is not None else "Mostra solo righe con problemi",
                   on_click=toggle_problem_rows)
    actions.button("Mostra le colonne predefinite" if ss.problem_cols_only else "Mostra solo colonne con problemi",
                   on_click=toggle_problem_columns)
    actions.button("Accetta come vuoti i valori non validi", on_click=accept_invalid_as_empty,
                   disabled=not current.conv_issues,
                   help="Svuota le celle con valori non interpretabili (es. 'n.a.' in un KPI)")
    actions.download_button("Scarica Excel (formato Metadata_ALL_DB)", batch_excel_bytes(current),
                            file_name="Metadata_ALL_DB.xlsx",
                            help="Per chi vuole ancora modificare in Excel: il file si reimporta "
                                 "con la sorgente 'Excel Metadata + KPI'")
    show_flash()

    all_columns = current.review_columns()
    # changing the visible columns commits the pending edits and re-creates the grid: a grid
    # reports only the edits of its visible columns and stops editing once its columns change
    if ss.get("review_cols") is None:
        ss.review_cols = default_columns(current, issues)
    ss.review_cols = [c for c in ss.review_cols if c in all_columns]
    chosen = st.multiselect("Colonne visualizzate", all_columns, key="review_cols", on_change=commit_edits)
    # the grid gets the committed rows, never the pending edits: when its data changes the
    # grid drops every edit equal to the new cell value, so feeding the edits back would erase them
    display = display_frame(ss.batch, schema, problems_text(issues), view_rows(), settings.id_column)
    extra = [SOURCE_COL] if SOURCE_COL in display.columns else []
    config = column_config(current)
    st.data_editor(display, key=editor_key(), num_rows="fixed", column_config=config,
                   column_order=[PROBLEMS_COL, *[c for c in all_columns if c in chosen], *extra],
                   height=min(600, 38 + 35 * len(display)), placeholder="—")
    inferred, groups = inferred_columns(current), inferred_groups(current)
    if inferred and groups:
        with st.expander("Valori dedotti automaticamente (modificabili per file)", expanded=True):
            st.caption(f"{len(groups)} file. Valori dedotti dal nome file o dal template per riga: una modifica "
                       "qui vale per tutte le righe (le ripetizioni) dello stesso file.")
            inferred_config = {c: v for c, v in config.items() if c in inferred}
            inferred_config.update({
                FILE_COL: st.column_config.TextColumn(FILE_COL, disabled=True, width="large"),
                ROWS_COL: st.column_config.TextColumn(ROWS_COL, disabled=True, help="righe della tabella sopra"),
                FROM_COL: st.column_config.TextColumn(FROM_COL, disabled=True),
                MIXED_COL: st.column_config.TextColumn(MIXED_COL, disabled=True,
                                                       help="colonne con valori diversi tra le righe del file: "
                                                            "qui compare quello della prima riga"),
            })
            frame = inferred_frame(current, schema, inferred, groups)
            st.data_editor(frame, key=inferred_key(), num_rows="fixed", hide_index=True,
                           column_config=inferred_config, on_change=apply_inferred_edits,
                           height=min(400, 38 + 35 * len(frame)), placeholder="—")
    missing = missing_support_values(issues, schema, options())
    if missing:
        with st.container(border=True):
            st.markdown("**Valori non presenti nei fogli di supporto del DB**")
            st.caption("Aggiungili qui al foglio: le righe vengono ricontrollate subito, senza tornare al passo 3.")
            for (column, value), rows in missing.items():
                field = schema.get(column)
                line = st.container(horizontal=True, vertical_alignment="center")
                line.markdown(f"`{column}` = `{value}` · righe {', '.join(str(r + 1) for r in rows)}")
                with line:
                    add_value_popover(field, label=f"Aggiungi al foglio {field.choices_sheet}", value=value,
                                      scope=f"review::{value}")
    if issues:
        with st.expander(f"Elenco problemi ({len(issues)})", expanded=n_err > 0):
            st.dataframe(issues_frame(issues), hide_index=True)
    nav(back=2, forward=4, label="Avanti: controlli e scrittura →", on_forward=lambda: (commit_edits(), goto(4)),
        on_back=lambda: (commit_edits(), goto(2)))


def edit_changes() -> tuple[dict, list[str]]:
    """(values to write, fields filled in the DB but empty in the form)."""
    record, values = ss.edit_record, ss.edit_values
    cleared = [c for c in schema.columns if c not in values and not is_missing(record.get(c))
               and schema.get(c).type != "id"]
    new = dict(values)
    if ss.get("clear_empty_fields"):
        new.update({c: None for c in cleared})
    diff = diff_frame(schema, record, new)
    return {row["Colonna"]: new.get(row["Colonna"]) for _, row in diff.iterrows()}, cleared


def review_edit() -> None:
    record, values = ss.edit_record, ss.edit_values
    if record is None or values is None:
        st.info("Carica un ID al passo 2 e compila il form al passo 3.")
        nav(back=1)
        return
    st.subheader(f"Modifiche all'ID {to_text(record.get(settings.id_column))}")
    changes, cleared = edit_changes()
    if cleared:
        st.checkbox(f"Svuota nel DB i campi lasciati vuoti nel form ({', '.join(cleared)})", key="clear_empty_fields")
        changes, _ = edit_changes()
    if not changes:
        st.info("Nessuna differenza rispetto al DB.")
    else:
        st.dataframe(diff_frame(schema, record, {**values, **changes}), hide_index=True)
    issues = validate_batch(batch_manual(schema, values), schema, settings, options())
    for issue in issues:
        (st.error if issue.severity == ERROR else st.warning)(f"{issue.column}: {issue.message}")
    nav(back=2, forward=4, disabled=not changes or any(i.severity == ERROR for i in issues))


# --- step 5: write ---------------------------------------------------------------------------

def reset_all() -> None:
    for key in ("base", "batch", "view_rows", "edit_record", "edit_values", "result", "source_report",
                "review_cols"):
        ss[key] = None
    ss.problem_cols_only = False
    ss.editor_version += 1
    goto(1)


def set_all_decisions(value: str, names: list[str]) -> None:
    for name in names:
        ss[f"kpi::{name}"] = value


def decisions_from_state(unknown: list[str]) -> dict[str, KpiDecision]:
    decisions = {}
    for name in unknown:
        choice = ss.get(f"kpi::{name}", CHOOSE)
        if choice == ADD:
            decisions[name] = KpiDecision("add")
        elif choice == DISCARD:
            decisions[name] = KpiDecision("discard")
        elif choice.startswith(MAP):
            decisions[name] = KpiDecision("map", choice[len(MAP):])
    return decisions


def show_result() -> None:
    result = ss.result
    if result is None:
        return
    if result.action == "append":
        st.success(f"Scritti gli ID {result.ids[0]}–{result.ids[-1]} in {result.elapsed_s:.1f} s.")
    else:
        st.success(f"ID {result.ids[0]} aggiornato in {result.elapsed_s:.1f} s.")
    for message in result.messages:
        st.write("• " + message)
    st.caption(f"Backup del DB prima della scrittura: {result.backup}")
    st.button("Nuovo inserimento", type="primary", on_click=reset_all)


def step_write() -> None:
    st.header("5 · Controlli e scrittura")
    if not require_db():
        return
    refresh_if_changed()
    show_flash()
    if ss.result is not None:
        show_result()
        return
    snap = snapshot()
    if ss.source == "edit":
        write_edit(snap)
        return
    if ss.batch is None:
        st.info("Componi prima le righe al passo 3.")
        nav(back=2)
        return
    batch = ss.batch
    blockers = [i for i in snap.issues if i.severity == ERROR]
    for issue in blockers:
        st.error(f"DB: {issue.message}")
    if snap.lock:
        st.error(f"{snap.lock}: chiudi il file prima di scrivere.")
    issues = [i for i in batch_issues(batch) if i.severity == ERROR]
    if issues:
        st.error(f"{len(issues)} errori nelle righe: correggili al passo 4.")
        st.dataframe(issues_frame(issues), hide_index=True)

    known, unknown, suggestions = kpi_resolution(batch, snap.kpi_names, settings)
    if unknown:
        st.subheader("KPI non presenti in TableKPI")
        st.caption("Scegli per ciascuno: nuova colonna, colonna esistente oppure scarta.")
        b = st.container(horizontal=True)
        b.button("Tutti: nuova colonna", on_click=set_all_decisions, args=(ADD, unknown))
        b.button("Tutti: scarta", on_click=set_all_decisions, args=(DISCARD, unknown))
        choices = [CHOOSE, ADD, DISCARD, *[MAP + k for k in snap.kpi_names]]
        for name in unknown:
            ss.setdefault(f"kpi::{name}", MAP + suggestions[name] if name in suggestions else CHOOSE)
            st.selectbox(name, choices, key=f"kpi::{name}")
    decisions = decisions_from_state(unknown)
    decision_issues = check_decisions(batch, snap.kpi_names, decisions, settings)
    for issue in decision_issues:
        st.warning(issue.message)

    ready = not blockers and not snap.lock and not issues and not decision_issues
    if ready:
        meta_rows, kpi_rows, new_columns, skipped = preview_write(snap, batch, schema, settings, decisions)
        first = snap.next_id
        st.info(f"Verranno aggiunte **{batch.n_rows} righe** a {settings.metadata_table} e "
                f"{settings.kpi_table} con **ID {first}–{first + batch.n_rows - 1}**"
                + (f"; nuove colonne KPI: {', '.join(new_columns)}" if new_columns else "") + ".")
        if skipped:
            st.warning("Colonne non presenti in TableMet (i valori non saranno scritti): " + ", ".join(skipped))
        with st.expander("Anteprima delle righe come verranno salvate"):
            filled = [c for c in meta_rows.columns if meta_rows[c].map(lambda v: not is_missing(v)).any()]
            st.caption(settings.metadata_table)
            st.dataframe(meta_rows[filled], hide_index=True)
            st.caption(settings.kpi_table)
            st.dataframe(kpi_rows, hide_index=True)
        st.caption(f"Prima di scrivere viene fatto un backup in {backup_root(settings)}")
    if st.button("Scrivi nel DB", type="primary", disabled=not ready):
        with st.status("Scrittura in corso…", expanded=True) as status:
            try:
                ss.result = write_batch(snap.path, batch, schema, settings, decisions)
                status.update(label="Scrittura completata", state="complete")
            except DBError as exc:
                status.update(label="Scrittura annullata: il DB non è stato modificato", state="error")
                st.error(str(exc))
                return
        connect(str(snap.path))
        st.rerun()
    nav(back=3)


def write_edit(snap) -> None:
    record, values = ss.edit_record, ss.edit_values
    if record is None or values is None:
        st.info("Carica un ID al passo 2 e compila il form al passo 3.")
        nav(back=1)
        return
    record_id = record.get(settings.id_column)
    changed, _ = edit_changes()
    st.dataframe(diff_frame(schema, record, {**values, **changed}), hide_index=True)
    if snap.lock:
        st.error(f"{snap.lock}: chiudi il file prima di scrivere.")
    st.caption("Vengono aggiornati anche IDName e Date della riga KPI con lo stesso ID. "
               f"Backup automatico in {backup_root(settings)}")
    if st.button(f"Sovrascrivi l'ID {to_text(record_id)}", type="primary",
                 disabled=not changed or bool(snap.lock) or not snap.writable):
        try:
            with st.spinner("Scrittura…"):
                ss.result = overwrite_metadata(snap.path, record_id, changed, schema, settings)
        except DBError as exc:
            st.error(str(exc))
            return
        connect(str(snap.path))
        st.rerun()
    nav(back=3)


# --- layout ----------------------------------------------------------------------------------

init_state()
keep_widget_values()
with st.sidebar:
    st.title("KPI metadata → DB")
    # leaving step 4 from here must not lose the edits made in the grid
    st.radio("Passi", STEPS, key="step", label_visibility="collapsed", on_change=commit_edits)
    snap = snapshot()
    st.divider()
    if snap is None:
        st.caption("DB non collegato")
    else:
        state = "✅ scrivibile" if snap.writable and not snap.lock else "⛔ non scrivibile"
        st.caption(f"**{snap.path.name}** · {state}")
        st.caption(f"{len(snap.metadata.df) if snap.metadata is not None else 0} righe · prossimo ID {snap.next_id}")
    st.divider()
    st.caption(f"Schema: {settings.schema_path.name} · date nel DB: "
               f"{'testo gg/mm/aaaa' if settings.date_storage == 'text' else 'date Excel'}")
    if st.button("Ricarica configurazione"):
        get_config.clear()
        st.rerun()

{STEPS[0]: step_db, STEPS[1]: step_source, STEPS[2]: step_form, STEPS[3]: step_review,
 STEPS[4]: step_write}[ss.step]()
