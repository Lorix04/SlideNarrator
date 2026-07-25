"""Elaborazione batch e ripresa sicura per Slide Narrator.

Manifest JSON:
{
  "defaults": {"voice":"it-IT-IsabellaNeural", "rate":"+0%"},
  "jobs": [
    {"input_pptx":"a.pptx", "scripts_xlsx":"a.xlsx", "output_pptx":"a_audio.pptx"},
    {"mode":"fix", "input_pptx":"b.pptx", "scripts_xlsx":"b.xlsx", "output_pptx":"b_FIX.pptx"}
  ]
}
"""
from __future__ import annotations
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import slide_narrator


def _completed(output: Path) -> bool:
    report = output.with_name(f"{output.stem}_esecuzione.json")
    if not output.exists() or not report.exists():
        return False
    try:
        return json.loads(report.read_text(encoding="utf-8")).get("status") == "completed"
    except Exception:
        return False


def run_manifest(manifest_path: str, *, force: bool = False, stop_on_error: bool = False) -> str:
    manifest_file = Path(manifest_path).resolve()
    data = json.loads(manifest_file.read_text(encoding="utf-8"))
    defaults = dict(data.get("defaults") or {})
    jobs = list(data.get("jobs") or [])
    if not jobs:
        raise ValueError("Il manifest non contiene lavori.")
    results = []
    for index, raw in enumerate(jobs, 1):
        job = {**defaults, **dict(raw)}
        for key in ("input_pptx", "output_pptx"):
            if key not in job:
                raise ValueError(f"Lavoro {index}: manca {key}")
        for key in ("input_pptx", "scripts_xlsx", "output_pptx"):
            if job.get(key) and not Path(job[key]).is_absolute():
                job[key] = str((manifest_file.parent / job[key]).resolve())
        output = Path(job["output_pptx"])
        if not force and _completed(output):
            results.append({"index": index, "status": "skipped_completed", "output": str(output)})
            print(f"[{index}/{len(jobs)}] già completato: {output.name}")
            continue
        mode = str(job.pop("mode", "generate")).lower()
        try:
            print(f"[{index}/{len(jobs)}] avvio {mode}: {Path(job['input_pptx']).name}")
            if mode == "fix":
                slide_narrator.process_fix(**job)
            else:
                slide_narrator.process(**job)
            results.append({"index": index, "status": "completed", "output": str(output)})
        except Exception as exc:
            results.append({"index": index, "status": "error", "error": str(exc), "output": str(output)})
            print(f"[{index}/{len(jobs)}] ERRORE: {exc}", file=sys.stderr)
            if stop_on_error:
                break
    summary = manifest_file.with_name(f"{manifest_file.stem}_risultati_{datetime.now():%Y%m%d_%H%M%S}.json")
    summary.write_text(json.dumps({"manifest": str(manifest_file), "results": results}, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(summary)


def main():
    p = argparse.ArgumentParser(description="Elabora più PowerPoint con ripresa dei lavori già completati.")
    p.add_argument("manifest", help="Manifest JSON dei lavori")
    p.add_argument("--force", action="store_true", help="Rielabora anche gli output già completati")
    p.add_argument("--stop-on-error", action="store_true", help="Interrompi al primo errore")
    args = p.parse_args()
    print(run_manifest(args.manifest, force=args.force, stop_on_error=args.stop_on_error))


if __name__ == "__main__":
    main()
