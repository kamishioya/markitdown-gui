from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QSettings, QTimer
from PySide6.QtWidgets import QApplication

from markitdown_gui._main_window import MainWindow


ROOT_DIR = Path(__file__).resolve().parents[3]
TEST_FILES_DIR = ROOT_DIR / "packages" / "markitdown" / "tests" / "test_files"
PDF_FILE = TEST_FILES_DIR / "test.pdf"
DOCX_FILE = TEST_FILES_DIR / "test.docx"
XLSX_FILE = TEST_FILES_DIR / "test.xlsx"


def _excel_com_available() -> bool:
    command = (
        "try { "
        "$excel = New-Object -ComObject Excel.Application; "
        "$excel.DisplayAlerts = $false; "
        "$excel.Quit(); "
        "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null; "
        "Write-Output 'OK' "
        "} catch { Write-Output 'NG' }"
    )
    completed = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            command,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    return completed.returncode == 0 and "OK" in completed.stdout


def main() -> int:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

    app = QApplication(sys.argv)

    with tempfile.TemporaryDirectory(prefix="markitdown-gui-smoke-") as temp_dir:
        output_dir = Path(temp_dir)
        settings = QSettings(
            str(output_dir / "smoke-settings.ini"),
            QSettings.Format.IniFormat,
        )
        settings.clear()
        window = MainWindow(settings=settings, default_copilot_command="")
        failures: list[str] = []
        include_xlsx = _excel_com_available()

        def finish(exit_code: int) -> None:
            print(f"SMOKE_EXIT={exit_code}")
            app.exit(exit_code)

        def verify() -> None:
            pdf_output = output_dir / "test.md"
            docx_output = output_dir / "test (1).md"
            xlsx_output = output_dir / "test (2).md" if include_xlsx else None

            if not pdf_output.exists():
                failures.append(f"Missing output: {pdf_output}")
            if not docx_output.exists():
                failures.append(f"Missing output: {docx_output}")
            if xlsx_output is not None and not xlsx_output.exists():
                failures.append(f"Missing output: {xlsx_output}")

            if pdf_output.exists():
                pdf_text = pdf_output.read_text(encoding="utf-8")
                if "While there is contemporaneous exploration of multi-agent approaches" not in pdf_text:
                    failures.append("PDF output did not contain the expected text.")

            if docx_output.exists():
                docx_text = docx_output.read_text(encoding="utf-8")
                if "AutoGen: Enabling Next-Gen LLM Applications via Multi-Agent Conversation" not in docx_text:
                    failures.append("DOCX output did not contain the expected text.")
                if "# Abstract" not in docx_text:
                    failures.append("DOCX output did not contain the expected heading.")

            if xlsx_output is not None and xlsx_output.exists():
                xlsx_text = xlsx_output.read_text(encoding="utf-8")
                if "Alpha" not in xlsx_text:
                    failures.append("XLSX output did not contain the expected workbook text.")
                if "Beta" not in xlsx_text:
                    failures.append("XLSX output did not contain the expected column heading.")

            if failures:
                for failure in failures:
                    print(f"SMOKE_FAIL: {failure}")
                finish(1)
                return

            print(f"SMOKE_OK: {pdf_output}")
            print(f"SMOKE_OK: {docx_output}")
            if xlsx_output is not None:
                print(f"SMOKE_OK: {xlsx_output}")
            else:
                print("SMOKE_SKIP: XLSX conversion skipped because Excel COM is unavailable.")
            finish(0)

        def start() -> None:
            window._open_output_checkbox.setChecked(False)
            window._overwrite_checkbox.setChecked(False)
            window._keep_data_uris_checkbox.setChecked(False)
            window._copilot_checkbox.setChecked(False)
            window._copilot_command_edit.clear()
            window._output_dir_edit.setText(str(output_dir))
            input_paths = [str(PDF_FILE), str(DOCX_FILE)]
            if include_xlsx:
                input_paths.append(str(XLSX_FILE))

            window._add_paths(input_paths)
            expected_rows = len(input_paths)
            if window._file_table.rowCount() != expected_rows:
                failures.append(
                    f"Expected {expected_rows} rows after file selection, found {window._file_table.rowCount()}."
                )
                verify()
                return

            window._start_conversion()
            if window._worker is None:
                failures.append("Conversion worker was not created.")
                verify()
                return

            window._worker.finished.connect(lambda *_args: QTimer.singleShot(0, verify))

        QTimer.singleShot(0, start)
        QTimer.singleShot(
            120000,
            lambda: (print("SMOKE_FAIL: Timeout waiting for GUI conversion."), finish(2)),
        )
        return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())