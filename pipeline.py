#!/usr/bin/env python3
"""End-to-end pipeline: Azure DI JSON -> cleaned fields -> rate matrix XLSX."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from build_matrix import build_matrix_workbook
from clean_di_json import write_cleaned_json
from file_selection import prompt_input_json
from postal_zones import postal_zones_txt_path_for_matrix
from project_paths import OUTPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs, is_colab_environment

_IN_COLAB = is_colab_environment()


@dataclass(frozen=True)
class PipelineResult:
    source_path: Path
    cleaned_path: Path
    matrix_path: Path
    zones_txt_path: Path
    zone_row_count: int
    field_block_count: int


def cleaned_json_path(source_path: Path) -> Path:
    return PROCESSING_DIR / f"{source_path.stem}.cleaned.json"


def run_clean_step(source_path: Path) -> tuple[Path, dict]:
    destination = cleaned_json_path(source_path)
    cleaned = write_cleaned_json(source_path, destination)
    return destination, cleaned


def run_matrix_step(cleaned_path: Path, output_path: Path | None = None) -> Path:
    if output_path is None:
        stem = cleaned_path.stem.replace(".cleaned", "")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = OUTPUT_DIR / f"matrix_{stem}_{timestamp}.xlsx"
    return build_matrix_workbook(cleaned_path, output_path=output_path)


def run_pipeline(source_path: Path | None = None) -> PipelineResult:
    ensure_workspace_dirs()

    if source_path is None:
        source_path = prompt_input_json()
    else:
        source_path = source_path.resolve()
        if not source_path.is_file():
            raise FileNotFoundError(f"Source file not found: {source_path}")

    print(f"\n[1/2] Cleaning Document Intelligence JSON...")
    print(f"      Source: {source_path.name}")
    cleaned_path, cleaned = run_clean_step(source_path)
    documents = cleaned.get("documents") or []
    fields = documents[0].get("fields", {}) if documents else {}
    field_block_count = len(fields)
    print(f"      Wrote {cleaned_path}")
    for name, value in fields.items():
        if isinstance(value, list):
            print(f"        {name}: {len(value)} row(s)")

    print("\n[2/2] Building matrix workbook...")
    matrix_path = run_matrix_step(cleaned_path)
    main_costs = fields.get("MainCosts") or []
    zone_row_count = len(
        [key for key in (main_costs[0] if main_costs else {}) if str(key).startswith("Zone")]
    )
    print(f"      Wrote {matrix_path}")
    zones_txt_path = postal_zones_txt_path_for_matrix(matrix_path)
    print(f"      Wrote {zones_txt_path}")
    print(f"        Rates rows (zones): {zone_row_count}")

    return PipelineResult(
        source_path=source_path,
        cleaned_path=cleaned_path,
        matrix_path=matrix_path,
        zones_txt_path=zones_txt_path,
        zone_row_count=zone_row_count,
        field_block_count=field_block_count,
    )


def print_summary(result: PipelineResult) -> None:
    print("\nPipeline complete.")
    print(f"  Source JSON:     {result.source_path}")
    print(f"  Cleaned JSON:    {result.cleaned_path}")
    print(f"  Matrix XLSX:     {result.matrix_path}")
    print(f"  Postal zones:    {result.zones_txt_path}")
    print(f"  Zone lanes:      {result.zone_row_count}")
    print(f"  Field blocks:    {result.field_block_count}")


def _strip_jupyter_args(argv: list[str]) -> list[str]:
    """Drop ``-f connection.json`` and other kernel flags from notebook argv."""
    cleaned: list[str] = []
    index = 0
    while index < len(argv):
        arg = argv[index]
        if arg in ("-f", "--f"):
            index += 2
            continue
        if arg.startswith("-") and arg not in ("-",):
            index += 1
            continue
        cleaned.append(arg)
        index += 1
    return cleaned


def _default_cli_argv() -> list[str]:
    """Ignore Jupyter/Colab kernel flags such as ``-f connection.json``."""
    if _IN_COLAB:
        return []
    try:
        from IPython import get_ipython

        shell = get_ipython()
        if shell is not None and shell.__class__.__name__ in (
            "ZMQInteractiveShell",
            "Shell",
        ):
            return []
    except ImportError:
        pass
    return _strip_jupyter_args(sys.argv[1:])


def colab_run(source_path: Path | None = None) -> PipelineResult:
    """Run the pipeline in Colab/notebooks without CLI argument parsing."""
    result = run_pipeline(source_path)
    print_summary(result)
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Run Dachser rate pipeline (clean JSON + matrix XLSX)."
    )
    parser.add_argument(
        "source",
        nargs="?",
        type=Path,
        help="Azure DI JSON in input/ (optional; prompts if omitted)",
    )
    args = parser.parse_args(argv if argv is not None else _default_cli_argv())
    result = run_pipeline(args.source)
    print_summary(result)
    return 0


if __name__ == "__main__":
    exit_code = main()
    if _IN_COLAB:
        pass
    else:
        try:
            from IPython import get_ipython

            if get_ipython() is None:
                raise SystemExit(exit_code)
        except ImportError:
            raise SystemExit(exit_code)
