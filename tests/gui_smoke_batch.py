from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from PySide6.QtCore import QSettings, QTimer
from PySide6.QtWidgets import QApplication

from markitdown_gui._main_window import MainWindow


PDF_TEXT = "MarkItDown GUI smoke test PDF sample"
DOCX_TITLE = "MarkItDown GUI smoke test document"
DOCX_BODY = "AutoGen enables next-generation LLM workflows through multi-agent conversation."
XLSX_HEADERS = ["Alpha", "Beta"]
XLSX_ROW = ["A-1", "B-1"]


def _write_sample_pdf(path: Path) -> None:
    escaped_text = (
        PDF_TEXT.replace("\\", "\\\\")
        .replace("(", "\\(")
        .replace(")", "\\)")
    )
    stream = (
        "BT\n"
        "/F1 12 Tf\n"
        "72 720 Td\n"
        f"({escaped_text}) Tj\n"
        "ET\n"
    )
    stream_bytes = stream.encode("ascii")
    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(stream_bytes)} >>\nstream\n{stream}endstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    chunks: list[bytes] = [b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"]
    offsets: list[int] = []
    current_offset = len(chunks[0])

    for index, obj in enumerate(objects, start=1):
        obj_bytes = f"{index} 0 obj\n{obj}\nendobj\n".encode("ascii")
        offsets.append(current_offset)
        chunks.append(obj_bytes)
        current_offset += len(obj_bytes)

    xref_offset = current_offset
    xref = [f"xref\n0 {len(objects) + 1}\n".encode("ascii")]
    xref.append(b"0000000000 65535 f \n")
    for offset in offsets:
        xref.append(f"{offset:010d} 00000 n \n".encode("ascii"))
    xref.append(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF\n"
        ).encode("ascii")
    )
    chunks.extend(xref)
    path.write_bytes(b"".join(chunks))


def _write_sample_docx(path: Path) -> None:
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>{escape(DOCX_TITLE)}</w:t></w:r></w:p>
    <w:p><w:r><w:t>{escape(DOCX_BODY)}</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'''
    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
'''
    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'''

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", rels_xml)
        archive.writestr("word/document.xml", document_xml)


def _write_sample_xlsx(path: Path) -> None:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Smoke"
    sheet.append(XLSX_HEADERS)
    sheet.append(XLSX_ROW)
    workbook.save(path)


def _create_input_files(input_dir: Path) -> tuple[Path, Path, Path]:
    pdf_file = input_dir / "test.pdf"
    docx_file = input_dir / "test.docx"
    xlsx_file = input_dir / "test.xlsx"
    _write_sample_pdf(pdf_file)
    _write_sample_docx(docx_file)
    _write_sample_xlsx(xlsx_file)
    return pdf_file, docx_file, xlsx_file


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
        input_dir = output_dir / "inputs"
        input_dir.mkdir(parents=True, exist_ok=True)
        pdf_file, docx_file, xlsx_file = _create_input_files(input_dir)
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
                if PDF_TEXT not in pdf_text:
                    failures.append("PDF output did not contain the expected text.")

            if docx_output.exists():
                docx_text = docx_output.read_text(encoding="utf-8")
                if DOCX_TITLE not in docx_text:
                    failures.append("DOCX output did not contain the expected text.")
                if DOCX_BODY not in docx_text:
                    failures.append("DOCX output did not contain the expected body text.")

            if xlsx_output is not None and xlsx_output.exists():
                xlsx_text = xlsx_output.read_text(encoding="utf-8")
                if XLSX_HEADERS[0] not in xlsx_text:
                    failures.append("XLSX output did not contain the expected workbook text.")
                if XLSX_HEADERS[1] not in xlsx_text:
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
            input_paths = [str(pdf_file), str(docx_file)]
            if include_xlsx:
                input_paths.append(str(xlsx_file))

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