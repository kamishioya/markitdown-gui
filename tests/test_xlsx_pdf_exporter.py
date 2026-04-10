from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from markitdown_gui._xlsx_pdf_exporter import default_xlsx_pdf_exporter


class XlsxPdfExporterTests(unittest.TestCase):
    def test_exporter_uses_hidden_process_flags(self) -> None:
        completed = subprocess.CompletedProcess(
            args=["powershell"],
            returncode=0,
            stdout="",
            stderr="",
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.xlsx"
            pdf_path = temp_root / "sample.pdf"
            source_path.write_bytes(b"PK\x03\x04")

            def run_side_effect(*args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
                pdf_path.write_bytes(b"%PDF-1.4")
                return completed

            with mock.patch(
                "markitdown_gui._xlsx_pdf_exporter.detect_powershell_command",
                return_value=r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
            ):
                with mock.patch(
                    "markitdown_gui._xlsx_pdf_exporter.subprocess.run",
                    side_effect=run_side_effect,
                ) as run_mock:
                    default_xlsx_pdf_exporter(source_path, pdf_path)

        if os.name == "nt":
            self.assertEqual(
                run_mock.call_args.kwargs.get("creationflags"),
                getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        else:
            self.assertNotIn("creationflags", run_mock.call_args.kwargs)


if __name__ == "__main__":
    unittest.main()