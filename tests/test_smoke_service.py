import pandas as pd

import core
from core.app_service import DiplomaGenerationService


class _FakeExcel:
    def __init__(self, sheet_names):
        self.sheet_names = sheet_names

    def parse(self, _sheet_name, header=None):
        del header
        return pd.DataFrame([[1]] * 8)


def _prepare_preflight_monkeypatch(monkeypatch, group: str):
    import core.app_service as app_service
    prefix = DiplomaGenerationService._target_sheet_prefix(group)

    monkeypatch.setattr(DiplomaGenerationService, "validate_inputs", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(app_service, "assert_program_subject_mapping", lambda _group: None)
    monkeypatch.setattr(
        app_service,
        "validate_program_subject_mapping",
        lambda _group: {"extra_count": 0, "extra_examples": []},
    )
    monkeypatch.setattr(app_service, "get_config", lambda _group, _lang: ({}, {}, "template.xlsx"))
    monkeypatch.setattr(app_service.os.path, "exists", lambda _path: True)
    monkeypatch.setattr(
        DiplomaGenerationService,
        "_load_excel",
        staticmethod(lambda _source: _FakeExcel([f"{prefix}-1", f"{prefix}-2"])),
    )


def test_core_lazy_export_service_available():
    assert hasattr(core, "DiplomaGenerationService")
    service = core.DiplomaGenerationService()
    assert isinstance(service, DiplomaGenerationService)


def test_preflight_smoke_3f(monkeypatch, tmp_path):
    _prepare_preflight_monkeypatch(monkeypatch, "3F")
    service = DiplomaGenerationService()
    prefix = service._target_sheet_prefix("3F")

    report = service.preflight_checks(
        source_file="dummy.xlsx",
        group="3F",
        lang="ALL",
        output_dir=str(tmp_path / "out_3f"),
    )

    assert report["ok"] is True
    assert report["errors"] == []
    assert report["sheet_count"] == 2
    assert report["target_sheets"] == [f"{prefix}-1", f"{prefix}-2"]


def test_preflight_smoke_3d(monkeypatch, tmp_path):
    _prepare_preflight_monkeypatch(monkeypatch, "3D")
    service = DiplomaGenerationService()

    report = service.preflight_checks(
        source_file="dummy.xlsx",
        group="3D",
        lang="ALL",
        output_dir=str(tmp_path / "out_3d"),
    )

    assert report["ok"] is True
    assert report["errors"] == []
    assert report["sheet_count"] == 2
    assert report["target_sheets"] == ["3D-1", "3D-2"]


def test_generate_batch_basic_service_run(monkeypatch, tmp_path):
    import core.app_service as app_service

    class _FakeGenerator:
        def __init__(self):
            self.closed = False
            self.filled = 0

        def fill_student_data(self, _student, _structured, lang="kz"):
            del lang
            self.filled += 1

        def close(self):
            self.closed = True

    monkeypatch.setattr(
        DiplomaGenerationService,
        "preflight_checks",
        lambda *_args, **_kwargs: {
            "ok": True,
            "errors": [],
            "warnings": [],
            "target_sheets": ["3F-1"],
        },
    )
    monkeypatch.setattr(
        DiplomaGenerationService,
        "_load_excel",
        staticmethod(lambda _source: _FakeExcel(["3F-1"])),
    )
    monkeypatch.setattr(
        app_service,
        "parse_excel_sheet",
        lambda _df, _sheet_name, start_row=5: [{"name": "Student One", "grades": []}],
    )
    monkeypatch.setattr(app_service, "build_diploma_pages", lambda _grades, _group: {1: []})
    monkeypatch.setattr(app_service, "get_config", lambda _group, _lang: ({}, {}, "template.xlsx"))
    monkeypatch.setattr(DiplomaGenerationService, "_make_generator", staticmethod(lambda *_args, **_kwargs: _FakeGenerator()))

    service = DiplomaGenerationService()
    result = service.generate_batch(
        source_file="dummy.xlsx",
        group="3F",
        lang="ALL",
        output_dir=str(tmp_path / "out"),
    )

    assert result["generated_count"] == 2
    assert result["error_count"] == 0
    assert result["sheets"]["3F-1"]["generated"] == 2