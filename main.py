import os
import argparse
import sys
from typing import Dict, Any

from core.app_service import DiplomaGenerationService
from core.exceptions import ValidationError

# Force utf-8 for Windows console
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def main():
    parser = argparse.ArgumentParser(description="Diploma Generator (Modular)")
    parser.add_argument("--source", type=str, required=True, help="Путь к исходному Excel-файлу")
    parser.add_argument("--group", type=str, required=True, choices=["3F", "3D"], help="Группа: 3F (IT) или 3D (Бухгалтеры)")
    parser.add_argument("--lang", type=str, default="ALL", choices=["KZ", "RU", "ALL"], help="Язык диплома")
    args = parser.parse_args()

    service = DiplomaGenerationService()

    def _progress(event: Dict[str, Any]) -> None:
        event_type = event.get("event")

        if event_type == "sheet_start":
            print(f"\nProcessing sheet: {event.get('sheet')}", flush=True)
            print(f"  Found {event.get('students', 0)} students.")
            return

        if event_type == "student_generated":
            out_path = event.get("output", "")
            print(f"    + {os.path.basename(out_path)}", flush=True)
            return

        if event_type == "student_error":
            out_path = event.get("output", "")
            print(f"    - [ERROR] Failed to generate {os.path.basename(out_path)}: {event.get('error')}")
            return

        if event_type == "sheet_done":
            print(
                f"  Done: generated={event.get('generated', 0)}, errors={event.get('errors', 0)}",
                flush=True,
            )
            return

        if event_type == "batch_done":
            print(
                f"\nGeneration finished: generated={event.get('generated_count', 0)}, "
                f"errors={event.get('error_count', 0)}",
                flush=True,
            )

    try:
        preflight = service.preflight_checks(args.source, args.group, args.lang, "output")
    except Exception as exc:
        print(f"Preflight failed: {exc}")
        return

    if not preflight["ok"]:
        for err in preflight["errors"]:
            print(f"[ERROR] {err}")
        return

    for warning in preflight.get("warnings", []):
        print(f"[WARN] {warning}")

    try:
        service.generate_batch(
            source_file=args.source,
            group=args.group,
            lang=args.lang,
            output_dir="output",
            progress_callback=_progress,
        )
    except ValidationError as exc:
        print(f"[ERROR] {exc}")
    except Exception as exc:
        print(f"[ERROR] Unexpected failure: {exc}")

if __name__ == "__main__":
    main()
