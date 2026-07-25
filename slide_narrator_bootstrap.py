from __future__ import annotations

import os
import runpy
import site
import sys
from pathlib import Path


def main() -> int:
    site_packages = os.environ.get("SLIDENARRATOR_SITE_PACKAGES", "").strip()
    if site_packages:
        path = Path(site_packages)
        if not path.is_dir():
            raise SystemExit(f"Runtime Slide Narrator non trovato: {path}")
        # addsitedir, a differenza del solo PYTHONPATH, elabora anche i file .pth.
        site.addsitedir(str(path))

    if len(sys.argv) < 2:
        raise SystemExit("Uso: slide_narrator_bootstrap.py <script.py> [argomenti]")

    target = sys.argv[1]
    if target == "-m":
        if len(sys.argv) < 3:
            raise SystemExit("Manca il nome del modulo dopo -m")
        module = sys.argv[2]
        sys.argv = [module, *sys.argv[3:]]
        runpy.run_module(module, run_name="__main__", alter_sys=True)
        return 0

    target_path = Path(target)
    if not target_path.is_absolute():
        target_path = Path.cwd() / target_path
    if not target_path.is_file():
        raise SystemExit(f"Script non trovato: {target_path}")
    sys.argv = [str(target_path), *sys.argv[2:]]
    runpy.run_path(str(target_path), run_name="__main__")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
