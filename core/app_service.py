import os
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from configs import get_config
from core.bridge import build_diploma_pages
from core.consistency import assert_program_subject_mapping, validate_program_subject_mapping
from core.exceptions import ConfigurationError, GenerationError, ParseError, ValidationError
from src.generator import DiplomaGenerator
from src.generator_3d import DiplomaGenerator3D
from src.generator_it import DiplomaGeneratorIT
from src.parser import parse_excel_sheet


ProgressCallback = Callable[[Dict[str, Any]], None]


class DiplomaGenerationService:
    """Application service that runs the existing generation pipeline."""

    VALID_GROUPS = {"3F", "3D"}
    VALID_LANGS = {"KZ", "RU", "ALL"}

    def validate_inputs(self, source_file: str, group: str, lang: str, output_dir: str) -> List[str]:
        errors: List[str] = []

        if not source_file:
            errors.append("Source file path is required.")
        elif not os.path.exists(source_file):
            errors.append(f"Source file not found: {source_file}")
        elif not source_file.lower().endswith((".xlsx", ".xlsm", ".xls")):
            errors.append("Source file must be an Excel file (.xlsx/.xlsm/.xls).")

        if group not in self.VALID_GROUPS:
            errors.append(f"Unsupported group: {group}. Allowed: {sorted(self.VALID_GROUPS)}")

        if lang not in self.VALID_LANGS:
            errors.append(f"Unsupported language: {lang}. Allowed: {sorted(self.VALID_LANGS)}")

        if not output_dir:
            errors.append("Output directory path is required.")

        return errors

    def preflight_checks(self, source_file: str, group: str, lang: str, output_dir: str) -> Dict[str, Any]:
        errors = self.validate_inputs(source_file, group, lang, output_dir)
        warnings: List[str] = []

        if errors:
            return {"ok": False, "errors": errors, "warnings": warnings, "report": None}

        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Cannot create output directory: {output_dir}. {exc}")
            return {"ok": False, "errors": errors, "warnings": warnings, "report": None}

        assert_program_subject_mapping(group)
        report = validate_program_subject_mapping(group)
        if report["extra_count"] > 0:
            warnings.append(
                f"Non-blocking mapping extras for {group}: {report['extra_count']} "
                f"(examples: {report['extra_examples'][:3]})"
            )

        langs_to_run = ["kz", "ru"] if lang == "ALL" else [lang.lower()]
        for one_lang in langs_to_run:
            _, _, template_name = get_config(group, one_lang)
            template_path = os.path.join("templates", template_name)
            if not os.path.exists(template_path):
                errors.append(f"Template not found: {template_path}")

        try:
            xl = self._load_excel(source_file)
        except ParseError as exc:
            errors.append(str(exc))
            return {"ok": False, "errors": errors, "warnings": warnings, "report": report}

        target_sheets = self._resolve_target_sheets(xl.sheet_names, group)
        if not target_sheets:
            prefix = self._target_sheet_prefix(group)
            errors.append(f"No sheets found starting with {prefix}")
            return {"ok": False, "errors": errors, "warnings": warnings, "report": report}

        # Basic structure check: first target sheet must have enough rows for student parsing.
        probe_df = xl.parse(target_sheets[0], header=None)
        if probe_df.shape[0] <= 5:
            errors.append(
                "Excel sheet structure is invalid: not enough rows for student parsing (expected > 5)."
            )

        return {
            "ok": len(errors) == 0,
            "errors": errors,
            "warnings": warnings,
            "report": report,
            "sheet_count": len(target_sheets),
            "target_sheets": target_sheets,
        }

    def generate_batch(
        self,
        source_file: str,
        group: str,
        lang: str = "ALL",
        output_dir: str = "output",
        progress_callback: Optional[ProgressCallback] = None,
    ) -> Dict[str, Any]:
        preflight = self.preflight_checks(source_file, group, lang, output_dir)
        if not preflight["ok"]:
            raise ValidationError("; ".join(preflight["errors"]))

        xl = self._load_excel(source_file)
        target_sheets = preflight["target_sheets"]
        langs_to_run = ["kz", "ru"] if lang == "ALL" else [lang.lower()]

        results: Dict[str, Any] = {
            "group": group,
            "lang": lang,
            "source": source_file,
            "output_dir": output_dir,
            "warnings": preflight["warnings"],
            "generated_count": 0,
            "error_count": 0,
            "errors": [],
            "sheets": {},
        }

        for sheet_name in target_sheets:
            df = xl.parse(sheet_name, header=None)
            df = df.iloc[:200, :300]
            students = parse_excel_sheet(df, sheet_name, start_row=5)

            sheet_result = {
                "students_found": len(students),
                "generated": 0,
                "errors": [],
            }
            results["sheets"][sheet_name] = sheet_result

            self._emit(progress_callback, {
                "event": "sheet_start",
                "sheet": sheet_name,
                "students": len(students),
            })

            for one_lang in langs_to_run:
                config, terms, template_name = get_config(group, one_lang)
                template_path = os.path.join("templates", template_name)

                for student in students:
                    safe_name = student["name"].replace("/", " ").replace("\\", " ")
                    out_name = f"{sheet_name}_{safe_name}_{one_lang.upper()}.xlsx"
                    out_path = os.path.join(output_dir, out_name)

                    generator = None
                    try:
                        structured = build_diploma_pages(student["grades"], group)
                        generator = self._make_generator(group, template_path, out_path, config, terms)
                        self._fill_student(group, generator, student, structured, one_lang)

                        results["generated_count"] += 1
                        sheet_result["generated"] += 1
                        self._emit(progress_callback, {
                            "event": "student_generated",
                            "sheet": sheet_name,
                            "student": student["name"],
                            "lang": one_lang,
                            "output": out_path,
                        })
                    except Exception as exc:  # noqa: BLE001
                        err = {
                            "sheet": sheet_name,
                            "student": student.get("name", "<unknown>"),
                            "lang": one_lang,
                            "output": out_path,
                            "error": str(exc),
                        }
                        results["error_count"] += 1
                        results["errors"].append(err)
                        sheet_result["errors"].append(err)
                        self._emit(progress_callback, {
                            "event": "student_error",
                            **err,
                        })
                    finally:
                        if generator is not None:
                            try:
                                generator.close()
                            except Exception:
                                pass
                            generator = None

            self._emit(progress_callback, {
                "event": "sheet_done",
                "sheet": sheet_name,
                "generated": sheet_result["generated"],
                "errors": len(sheet_result["errors"]),
            })

        self._emit(progress_callback, {
            "event": "batch_done",
            "generated_count": results["generated_count"],
            "error_count": results["error_count"],
            "output_dir": output_dir,
        })

        return results

    @staticmethod
    def _emit(progress_callback: Optional[ProgressCallback], payload: Dict[str, Any]) -> None:
        if progress_callback is None:
            return
        progress_callback(payload)

    @staticmethod
    def _target_sheet_prefix(group: str) -> str:
        return "3Ғ" if group == "3F" else group

    def _resolve_target_sheets(self, sheet_names: List[str], group: str) -> List[str]:
        prefix = self._target_sheet_prefix(group)
        return [name for name in sheet_names if name.startswith(prefix)]

    @staticmethod
    def _load_excel(source_file: str) -> pd.ExcelFile:
        try:
            return pd.ExcelFile(source_file, engine="calamine")
        except Exception:
            try:
                return pd.ExcelFile(source_file)
            except Exception as exc:  # noqa: BLE001
                raise ParseError(f"Failed to load Excel file: {source_file}. {exc}") from exc

    @staticmethod
    def _make_generator(
        group: str,
        template_path: str,
        out_path: str,
        config: Dict[str, Any],
        terms: Dict[str, str],
    ) -> Any:
        if group == "3D":
            return DiplomaGenerator3D(template_path, out_path, terms)
        if group == "3F":
            return DiplomaGeneratorIT(template_path, out_path, terms)
        return DiplomaGenerator(template_path, out_path, config, terms)

    @staticmethod
    def _fill_student(group: str, generator: Any, student: Dict[str, Any], structured: Dict[int, list], lang: str) -> None:
        try:
            if group in {"3D", "3F"}:
                generator.fill_student_data(student, structured, lang=lang)
            else:
                generator.fill_from_pages(student, structured, lang=lang)
        except Exception as exc:  # noqa: BLE001
            raise GenerationError(str(exc)) from exc
