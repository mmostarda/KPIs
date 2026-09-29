"""Command line: python -m kpimeta <command> ... (use -h for help)."""

from __future__ import annotations

import argparse
import glob
import sys
from pathlib import Path

from .batch import (Batch, BatchError, batch_from_excel, batch_from_frame, compose, export_batch_excel,
                    read_excel_table)
from .config import ConfigError, Settings, load_settings
from .excel_db import DBError, DBSnapshot, date_report, preview_write, read_db, write_batch
from .issues import ERROR, INFO, WARNING, Issue
from .lola import LolaError, export_all_db, is_lola_output, read_lola_files
from .mapping import load_mapping, save_mapping_csv
from .schema import Schema, load_schema
from .validate import (KpiDecision, canonicalize_choices, check_decisions, duplicates_vs_db,
                       internal_duplicates, kpi_resolution, summarize, validate_batch)
from .values import is_missing, to_text


def _print_issues(issues: list[Issue], limit: int = 60) -> None:
    order = {ERROR: 0, WARNING: 1, INFO: 2}
    for issue in sorted(issues, key=lambda i: (order.get(i.severity, 3), i.row or 0))[:limit]:
        print("  " + issue.label())
    if len(issues) > limit:
        print(f"  ... altri {len(issues) - limit}")


def _load(args) -> tuple[Settings, Schema]:
    settings = load_settings(args.config)
    return settings, load_schema(settings.schema_path)


def cmd_demo(args) -> int:
    from .demo import make_demo

    paths = make_demo(args.folder, args.config)
    print("Dati DEMO sintetici creati:")
    print(f"  DB:      {paths['db']}")
    print(f"  LoLa:    {', '.join(str(p) for p in paths['lola'])}")
    print(f"  mapping: {paths['mapping']}")
    return 0


def cmd_convert_mapping(args) -> int:
    mapping = load_mapping(args.source)
    target = args.output or Path(args.source).with_suffix(".csv")
    save_mapping_csv(mapping, target)
    print(f"{len(mapping.entries)} coppie LoLa -> DB scritte in {target}"
          + (f" ({mapping.skipped_rows} righe vuote/NA saltate)" if mapping.skipped_rows else ""))
    return 0


def cmd_doctor(args) -> int:
    settings, schema = _load(args)
    snapshot = read_db(args.db, schema, settings)
    print(f"DB: {snapshot.path}")
    for info in (snapshot.metadata, snapshot.kpis):
        if info is not None:
            print(f"  {info.name} (foglio {info.sheet}): {len(info.df)} righe, {len(info.headers)} colonne, "
                  f"range {info.ref}")
    print(f"  prossimo ID: {snapshot.next_id}")
    print(f"  colonne KPI: {len(snapshot.kpi_names)}")
    print(f"  template: {0 if snapshot.templates is None else len(snapshot.templates)}")
    report = date_report(snapshot, schema)
    if not report.empty:
        print("Come sono salvate oggi le colonne data:")
        print("  " + report.to_string(index=False).replace("\n", "\n  "))
    if snapshot.issues:
        print("Controlli:")
        _print_issues(snapshot.issues)
    print("Scrittura possibile." if snapshot.writable else "Scrittura NON possibile finché ci sono errori.")
    return 0 if snapshot.writable else 1


def _mapping(args, settings: Settings):
    return load_mapping(args.mapping or settings.lola_mapping_path)


def _lola_paths(patterns: list[str]) -> list[str]:
    """File arguments with wildcards expanded (the Windows shell does not do it); the
    intermediate files (ALL_DB.xlsx, ...) matched by a wildcard are skipped."""
    paths = []
    for pattern in patterns:
        if any(ch in pattern for ch in "*?["):
            matches = [p for p in sorted(glob.glob(pattern)) if is_lola_output(p)]
            if not matches:
                raise LolaError(f"nessun file corrisponde a {pattern}")
            paths += matches
        else:
            paths.append(pattern)
    return paths


def cmd_lola(args) -> int:
    settings, _ = _load(args)
    args.files = _lola_paths(args.files)
    frame, reports = read_lola_files(args.files, _mapping(args, settings), settings)
    for report in reports:
        print(f"{Path(report.path).name}: {report.data_rows} righe"
              + (f", {report.dropped_without_filename} senza File Name scartate" if report.dropped_without_filename else ""))
        if report.lola_names_missing:
            print(f"  colonne del mapping assenti: {', '.join(report.lola_names_missing)}")
        if report.unmapped_headers:
            print(f"  colonne ignorate (non nel mapping): {', '.join(report.unmapped_headers)}")
        for warning in report.warnings:
            print(f"  AVVISO: {warning}")
    output = args.output or Path(args.files[0]).with_name("ALL_DB.xlsx")
    export_all_db(frame, output)
    print(f"{len(frame)} righe scritte in {output}")
    return 0


def _common_values(args, snapshot: DBSnapshot, settings: Settings, schema: Schema) -> dict:
    values: dict = {}
    if args.template:
        templates = snapshot.templates
        if templates is None:
            raise BatchError("foglio dei template non disponibile")
        name_col = next((c for c in templates.columns
                         if str(c).lower() == settings.template_name_column.lower()), None)
        if name_col is None:
            raise BatchError(f"colonna {settings.template_name_column} non trovata in {settings.templates_sheet}")
        rows = templates[templates[name_col].map(lambda v: (to_text(v) or "").lower() == args.template.lower())]
        if rows.empty:
            raise BatchError(f"template '{args.template}' non trovato")
        values.update({schema.match(str(c)).column: v for c, v in rows.iloc[0].items()
                       if schema.match(str(c)) is not None and not is_missing(v)})
    if args.common:
        frame, _ = read_excel_table(args.common)
        if len(frame) != 1:
            raise BatchError(f"{args.common}: serve esattamente una riga (trovate {len(frame)})")
        values.update({schema.match(str(c)).column: v for c, v in frame.iloc[0].items()
                       if schema.match(str(c)) is not None and not is_missing(v)})
    for item in args.set or []:
        column, _, value = item.partition("=")
        field = schema.match(column.strip())
        if field is None:
            raise BatchError(f"--set: colonna '{column}' non presente nello schema")
        values[field.column] = value
    return values


def cmd_prepare(args) -> int:
    settings, schema = _load(args)
    snapshot = read_db(args.db, schema, settings)
    headers = snapshot.metadata.headers if snapshot.metadata else []
    if args.lola:
        frame, _ = read_lola_files(_lola_paths(args.lola), _mapping(args, settings), settings)
        base = batch_from_frame(frame, schema, settings, "lola", headers)
    else:
        base = batch_from_excel(args.excel, schema, settings, headers)
    batch = compose(base, schema, settings, _common_values(args, snapshot, settings, schema),
                    mode="fill" if args.fill else "overwrite", templates=snapshot.templates,
                    infer_template=args.infer_template, infer_filename=args.infer_filename)
    canonicalize_choices(batch, schema, snapshot.options(schema))
    issues = validate_batch(batch, schema, settings, snapshot.options(schema))
    export_batch_excel(batch, args.output)
    print(f"{batch.n_rows} righe, {len(batch.kpi_columns)} colonne KPI -> {args.output}")
    if issues:
        _print_issues(issues)
    return 0


def _checks(batch: Batch, snapshot: DBSnapshot, schema: Schema, settings: Settings,
            decisions: dict[str, KpiDecision]) -> list[Issue]:
    options = snapshot.options(schema)
    issues = list(snapshot.blockers)
    issues += validate_batch(batch, schema, settings, options)
    issues += internal_duplicates(batch, schema)
    if snapshot.metadata is not None:
        issues += duplicates_vs_db(batch, schema, snapshot.metadata.df)
    issues += check_decisions(batch, snapshot.kpi_names, decisions, settings)
    return issues


def cmd_write(args) -> int:
    settings, schema = _load(args)
    snapshot = read_db(args.db, schema, settings)
    batch = batch_from_excel(args.file, schema, settings, snapshot.metadata.headers if snapshot.metadata else [])
    canonicalize_choices(batch, schema, snapshot.options(schema))
    _, unknown, suggestions = kpi_resolution(batch, snapshot.kpi_names, settings)
    decisions: dict[str, KpiDecision] = {}
    for name in unknown:
        if args.new_kpi:
            decisions[name] = KpiDecision(args.new_kpi)
    for item in args.map or []:
        name, _, target = item.partition("=")
        decisions[name.strip()] = KpiDecision("map", target.strip())
    issues = _checks(batch, snapshot, schema, settings, decisions)
    counts = summarize(issues)
    print(f"{batch.n_rows} righe da scrivere, {len(batch.kpi_columns)} colonne KPI "
          f"({len(unknown)} non presenti nel DB{': ' + ', '.join(unknown) if unknown else ''})")
    for name, target in suggestions.items():
        print(f"  suggerimento: --map \"{name}={target}\"")
    if issues:
        _print_issues(issues)
    if counts[ERROR]:
        print(f"{counts[ERROR]} errori: niente è stato scritto.")
        return 1
    if args.dry_run or args.command == "check":
        meta, kpi, new_columns, skipped = preview_write(snapshot, batch, schema, settings, decisions)
        print(f"Controllo superato: verrebbero scritti gli ID {snapshot.next_id}-{snapshot.next_id + batch.n_rows - 1}"
              + (f", nuove colonne KPI: {', '.join(new_columns)}" if new_columns else ""))
        if skipped:
            print(f"  colonne non presenti in TableMet (non scritte): {', '.join(skipped)}")
        return 0
    if not args.yes:
        answer = input(f"Scrivere {batch.n_rows} righe in {snapshot.path.name}? [s/N] ").strip().lower()
        if answer not in ("s", "si", "sì", "y", "yes"):
            print("Annullato.")
            return 1
    result = write_batch(args.db, batch, schema, settings, decisions)
    print(f"Scritti gli ID {result.ids[0]}-{result.ids[-1]} in {result.elapsed_s:.1f} s. "
          f"Backup: {result.backup}")
    for message in result.messages:
        print(f"  {message}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="kpimeta", description="Metadati + KPI verso il DB Excel")
    parser.add_argument("--config", help="cartella di configurazione (default: python/config)")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("demo", help="crea dati DEMO sintetici per provare lo strumento")
    p.add_argument("folder")
    p.set_defaults(func=cmd_demo)

    p = sub.add_parser("convert-mapping", help="converte il mapping LoLa (.mat/.xlsx/.csv) in CSV")
    p.add_argument("source")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_convert_mapping)

    p = sub.add_parser("doctor", help="controlla il DB Excel senza modificarlo")
    p.add_argument("db")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("lola", help="legge output LoLa e crea ALL_DB.xlsx (come process_excel)")
    p.add_argument("files", nargs="+")
    p.add_argument("-m", "--mapping")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_lola)

    p = sub.add_parser("prepare", help="compone metadati + KPI in un Excel da rivedere")
    source = p.add_mutually_exclusive_group(required=True)
    source.add_argument("--lola", nargs="+", help="file di output LoLa")
    source.add_argument("--excel", help="Excel Metadata + KPI (es. Metadata_ALL_DB.xlsx)")
    p.add_argument("--db", required=True)
    p.add_argument("-m", "--mapping")
    p.add_argument("--template", help="valori comuni presi da Metadata_Templates (IDName_Template)")
    p.add_argument("--common", help="Excel con una riga di metadati comuni (export template)")
    p.add_argument("--set", action="append", metavar="COLONNA=VALORE")
    p.add_argument("--fill", action="store_true", help="i valori comuni riempiono solo le celle vuote")
    p.add_argument("--infer-filename", action="store_true")
    p.add_argument("--infer-template", action="store_true")
    p.add_argument("-o", "--output", required=True)
    p.set_defaults(func=cmd_prepare)

    for name, help_text in (("check", "controlla un Excel Metadata + KPI contro il DB (non scrive)"),
                            ("write", "scrive un Excel Metadata + KPI nel DB")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("file")
        p.add_argument("--db", required=True)
        p.add_argument("--new-kpi", choices=["add", "discard"],
                       help="cosa fare dei KPI non presenti in TableKPI")
        p.add_argument("--map", action="append", metavar="KPI=COLONNA_ESISTENTE")
        p.add_argument("--dry-run", action="store_true")
        p.add_argument("-y", "--yes", action="store_true", help="non chiedere conferma")
        p.set_defaults(func=cmd_write)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except (ConfigError, DBError, BatchError, LolaError, OSError) as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
