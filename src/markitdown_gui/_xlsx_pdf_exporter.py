from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from ._temp_cleanup import XLSX_PDF_SCRIPT_TEMP_PREFIX, build_temp_dir_prefix

XLSX_PDF_POWERSHELL_MISSING_MESSAGE = (
    "PowerShell is not available. XLSX to PDF conversion requires Windows PowerShell."
)
XLSX_PDF_EXCEL_MISSING_MESSAGE = (
    "Microsoft Excel is not available for XLSX to PDF conversion. Install Excel on Windows before converting .xlsx files."
)
XLSX_PDF_OUTPUT_MISSING_MESSAGE = "Excel did not generate the expected PDF output."
XLSX_PDF_TIMEOUT_MESSAGE = "XLSX to PDF conversion timed out."
XLSX_PDF_FAILURE_PREFIX = "XLSX to PDF conversion failed:"

_POWERSHELL_SCRIPT = """
param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$PdfPath
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $workbook = $excel.Workbooks.Open($SourcePath, 0, $true)
    $workbook.ExportAsFixedFormat(0, $PdfPath)
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
""".strip()


class XlsxPdfExportError(RuntimeError):
    pass


def _build_hidden_process_kwargs() -> dict[str, object]:
    if os.name != "nt":
        return {}
    return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}


def detect_powershell_command() -> str | None:
    configured_path = os.environ.get("MARKITDOWN_GUI_POWERSHELL", "").strip()
    if configured_path:
        return configured_path

    for command_name in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        command_path = shutil.which(command_name)
        if command_path:
            return command_path

    return None


def default_xlsx_pdf_exporter(source_path: Path, pdf_path: Path) -> None:
    powershell_command = detect_powershell_command()
    if not powershell_command:
        raise XlsxPdfExportError(XLSX_PDF_POWERSHELL_MISSING_MESSAGE)

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(
        prefix=build_temp_dir_prefix(XLSX_PDF_SCRIPT_TEMP_PREFIX)
    ) as temp_dir:
        script_path = Path(temp_dir) / "export_xlsx_to_pdf.ps1"
        script_path.write_text(_POWERSHELL_SCRIPT, encoding="utf-8")

        try:
            completed = subprocess.run(
                [
                    powershell_command,
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                    "-SourcePath",
                    str(source_path),
                    "-PdfPath",
                    str(pdf_path),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=120,
                check=False,
                **_build_hidden_process_kwargs(),
            )
        except FileNotFoundError as exc:
            raise XlsxPdfExportError(XLSX_PDF_POWERSHELL_MISSING_MESSAGE) from exc
        except subprocess.TimeoutExpired as exc:
            raise XlsxPdfExportError(XLSX_PDF_TIMEOUT_MESSAGE) from exc

    if completed.returncode != 0:
        raise XlsxPdfExportError(_build_failure_message(completed))

    if not pdf_path.exists():
        raise XlsxPdfExportError(XLSX_PDF_OUTPUT_MISSING_MESSAGE)


def _build_failure_message(completed: subprocess.CompletedProcess[str]) -> str:
    detail = _clean_text(f"{completed.stderr}\n{completed.stdout}")
    detail_lower = detail.lower()

    if (
        "excel.application" in detail_lower
        or "class not registered" in detail_lower
        or "retrieving the com class factory" in detail_lower
        or "activex component can't create object" in detail_lower
    ):
        return XLSX_PDF_EXCEL_MISSING_MESSAGE

    if not detail:
        return XLSX_PDF_FAILURE_PREFIX

    return f"{XLSX_PDF_FAILURE_PREFIX}\n{detail}"


def _clean_text(text: str) -> str:
    return text.replace("\r\n", "\n").strip()