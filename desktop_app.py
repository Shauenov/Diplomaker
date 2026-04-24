import json
import os
import sys
from datetime import datetime
from typing import Any, Dict

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.app_service import DiplomaGenerationService


class GenerationCancelledError(Exception):
    pass


class GenerationWorker(QThread):
    progress_event = Signal(dict)
    finished_ok = Signal(dict)
    finished_cancelled = Signal(str)
    finished_error = Signal(str)

    def __init__(self, source_file: str, group: str, lang: str, output_dir: str) -> None:
        super().__init__()
        self._source_file = source_file
        self._group = group
        self._lang = lang
        self._output_dir = output_dir
        self._cancel_requested = False

    def request_cancel(self) -> None:
        self._cancel_requested = True

    def run(self) -> None:
        service = DiplomaGenerationService()

        def _emit(payload: Dict[str, Any]) -> None:
            if self._cancel_requested or self.isInterruptionRequested():
                raise GenerationCancelledError("Generation cancelled by user")
            self.progress_event.emit(payload)

        try:
            result = service.generate_batch(
                source_file=self._source_file,
                group=self._group,
                lang=self._lang,
                output_dir=self._output_dir,
                progress_callback=_emit,
            )
            if self._cancel_requested or self.isInterruptionRequested():
                raise GenerationCancelledError("Generation cancelled by user")
            self.finished_ok.emit(result)
        except GenerationCancelledError as exc:
            self.finished_cancelled.emit(str(exc))
        except Exception as exc:  # noqa: BLE001
            self.finished_error.emit(str(exc))


class DiplomaDesktopWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.service = DiplomaGenerationService()
        self.worker: GenerationWorker | None = None
        self.log_file_path: str | None = None

        self.setWindowTitle("Diploma Generator Desktop")
        self.resize(860, 620)

        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget(self)
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        settings_box = QGroupBox("Generation Settings")
        settings_form = QFormLayout(settings_box)

        self.source_edit = QLineEdit()
        source_row = QHBoxLayout()
        source_row.addWidget(self.source_edit)
        source_btn = QPushButton("Browse")
        source_btn.clicked.connect(self._pick_source)
        source_row.addWidget(source_btn)
        settings_form.addRow("Source Excel:", self._wrap(source_row))

        self.group_combo = QComboBox()
        self.group_combo.addItems(["3F", "3D"])
        settings_form.addRow("Group:", self.group_combo)

        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["ALL", "KZ", "RU"])
        settings_form.addRow("Language:", self.lang_combo)

        self.output_edit = QLineEdit("output")
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_edit)
        output_btn = QPushButton("Browse")
        output_btn.clicked.connect(self._pick_output_dir)
        output_row.addWidget(output_btn)
        settings_form.addRow("Output Directory:", self._wrap(output_row))

        layout.addWidget(settings_box)

        controls = QHBoxLayout()
        self.preflight_btn = QPushButton("Validate")
        self.preflight_btn.clicked.connect(self._run_preflight)
        controls.addWidget(self.preflight_btn)

        self.generate_btn = QPushButton("Generate")
        self.generate_btn.clicked.connect(self._run_generation)
        controls.addWidget(self.generate_btn)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self._stop_generation)
        self.stop_btn.setEnabled(False)
        controls.addWidget(self.stop_btn)
        controls.addStretch(1)
        layout.addLayout(controls)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        layout.addWidget(self.log_view)

    @staticmethod
    def _wrap(inner_layout: QHBoxLayout) -> QWidget:
        wrapper = QWidget()
        wrapper.setLayout(inner_layout)
        return wrapper

    def _pick_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select source Excel file",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls)",
        )
        if path:
            self.source_edit.setText(path)

    def _pick_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select output directory", "")
        if path:
            self.output_edit.setText(path)

    def _current_inputs(self) -> Dict[str, str]:
        return {
            "source_file": self.source_edit.text().strip(),
            "group": self.group_combo.currentText(),
            "lang": self.lang_combo.currentText(),
            "output_dir": self.output_edit.text().strip() or "output",
        }

    def _append_log(self, message: str) -> None:
        self.log_view.appendPlainText(message)
        if self.log_file_path:
            with open(self.log_file_path, "a", encoding="utf-8") as log_file:
                log_file.write(f"{message}\n")

    def _start_log_file(self, output_dir: str) -> None:
        os.makedirs(output_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.log_file_path = os.path.join(output_dir, f"generation_{ts}.log")

    def _run_preflight(self) -> None:
        inputs = self._current_inputs()
        report = self.service.preflight_checks(**inputs)

        if report["ok"]:
            self.status_label.setText("Validation passed")
            self._append_log("[OK] Validation passed")
            if report.get("warnings"):
                for warning in report["warnings"]:
                    self._append_log(f"[WARN] {warning}")
            self._append_log(json.dumps({
                "sheet_count": report.get("sheet_count", 0),
                "target_sheets": report.get("target_sheets", []),
            }, ensure_ascii=False))
            return

        self.status_label.setText("Validation failed")
        for err in report.get("errors", []):
            self._append_log(f"[ERROR] {err}")
        QMessageBox.warning(self, "Validation", "Validation failed. See log panel.")

    def _run_generation(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(self, "Generation", "Generation is already running.")
            return

        inputs = self._current_inputs()
        preflight = self.service.preflight_checks(**inputs)
        if not preflight["ok"]:
            self.status_label.setText("Validation failed")
            for err in preflight["errors"]:
                self._append_log(f"[ERROR] {err}")
            QMessageBox.warning(self, "Generation", "Cannot start. Fix validation errors first.")
            return

        self._set_running_state(True)
        self._start_log_file(inputs["output_dir"])
        self.progress.setRange(0, 0)
        self.status_label.setText("Generation started")
        self._append_log("[INFO] Generation started")
        self._append_log(f"[INFO] Log file: {self.log_file_path}")

        self.worker = GenerationWorker(**inputs)
        self.worker.progress_event.connect(self._on_progress_event)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.finished_cancelled.connect(self._on_finished_cancelled)
        self.worker.finished_error.connect(self._on_finished_error)
        self.worker.start()

    def _stop_generation(self) -> None:
        if self.worker is None or not self.worker.isRunning():
            return
        self.worker.request_cancel()
        self.status_label.setText("Stopping...")
        self._append_log("[INFO] Stop requested by user")

    def _on_progress_event(self, payload: Dict[str, Any]) -> None:
        event_type = payload.get("event", "")
        if event_type == "sheet_start":
            self._append_log(f"[SHEET] {payload.get('sheet')} | students={payload.get('students', 0)}")
            return

        if event_type == "student_generated":
            student = payload.get("student", "")
            lang = str(payload.get("lang", "")).upper()
            self._append_log(f"[OK] {student} ({lang})")
            return

        if event_type == "student_error":
            student = payload.get("student", "")
            error = payload.get("error", "")
            self._append_log(f"[ERROR] {student}: {error}")
            return

        if event_type == "sheet_done":
            self._append_log(
                f"[SHEET DONE] {payload.get('sheet')} | generated={payload.get('generated', 0)} | "
                f"errors={payload.get('errors', 0)}"
            )
            return

        if event_type == "batch_done":
            self._append_log(
                f"[DONE] generated={payload.get('generated_count', 0)} | "
                f"errors={payload.get('error_count', 0)}"
            )

    def _on_finished_ok(self, result: Dict[str, Any]) -> None:
        self._set_running_state(False)
        self.progress.setRange(0, 1)
        self.progress.setValue(1)

        generated = result.get("generated_count", 0)
        errors = result.get("error_count", 0)
        self.status_label.setText(f"Done: generated={generated}, errors={errors}")
        self._append_log(f"[SUMMARY] output_dir={result.get('output_dir', '')}")
        QMessageBox.information(self, "Generation", f"Done: generated={generated}, errors={errors}")
        self.worker = None

    def _on_finished_cancelled(self, message: str) -> None:
        self._set_running_state(False)
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.status_label.setText("Generation cancelled")
        self._append_log(f"[CANCELLED] {message}")
        QMessageBox.information(self, "Generation", "Generation was cancelled.")
        self.worker = None

    def _on_finished_error(self, error_text: str) -> None:
        self._set_running_state(False)
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.status_label.setText("Generation failed")
        self._append_log(f"[FATAL] {error_text}")
        QMessageBox.critical(self, "Generation", error_text)
        self.worker = None

    def _set_running_state(self, running: bool) -> None:
        self.preflight_btn.setEnabled(not running)
        self.generate_btn.setEnabled(not running)
        self.stop_btn.setEnabled(running)


def run() -> int:
    app = QApplication(sys.argv)
    window = DiplomaDesktopWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(run())
