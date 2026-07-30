#!/usr/bin/env python3
"""CLI entry point for the Dachser end-to-end pipeline."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _code_dir_candidates() -> list[Path]:
    out: list[Path] = []
    try:
        out.append(Path(__file__).resolve().parent)
    except NameError:
        pass
    env = os.environ.get("DACHSER_CODE_DIR")
    if env:
        out.append(Path(env).expanduser())
    out.append(Path("/content/Ecolab-Dachser"))
    out.append(Path.cwd())
    return out


def _ensure_code_on_path() -> None:
    for root in _code_dir_candidates():
        root = root.resolve()
        if (root / "pipeline.py").is_file():
            root_str = str(root)
            if root_str not in sys.path:
                sys.path.insert(0, root_str)
            return
    raise RuntimeError(
        "Could not find Dachser code directory (need pipeline.py). "
        "Set DACHSER_CODE_DIR or run from /content/Ecolab-Dachser."
    )


_ensure_code_on_path()

from pipeline import colab_run, main
from project_paths import is_colab_environment

if __name__ == "__main__":
    if is_colab_environment():
        colab_run()
    else:
        raise SystemExit(main())
