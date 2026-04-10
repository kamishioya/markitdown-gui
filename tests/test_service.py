from __future__ import annotations

import tempfile
import unittest
from dataclasses import dataclass
from pathlib import Path

from markitdown_gui._service import ConversionOptions, MarkItDownService, RuntimeDependencyError
from markitdown_gui._xlsx_pdf_exporter import XlsxPdfExportError


@dataclass
class FakeResult:
    markdown: str
    title: str | None = None


class FakeConverter:
    def __init__(self) -> None:
        self.calls: list[tuple[str, dict[str, object]]] = []

    def convert(self, source: str, **kwargs: object) -> FakeResult:
        self.calls.append((source, kwargs))
        return FakeResult(markdown="# converted", title="Sample")


class MissingModuleConverter:
    def convert(self, source: str, **kwargs: object) -> FakeResult:
        raise ModuleNotFoundError("pdfplumber")


class FakeXlsxPdfExporter:
    def __init__(self) -> None:
        self.calls: list[tuple[Path, Path]] = []

    def __call__(self, source_path: Path, output_path: Path) -> None:
        self.calls.append((source_path, output_path))
        output_path.write_bytes(b"%PDF-1.4")


class MissingXlsxPdfExporter:
    def __call__(self, source_path: Path, output_path: Path) -> None:
        raise XlsxPdfExportError("Microsoft Excel is not available for XLSX to PDF conversion. Install Excel on Windows before converting .xlsx files.")


class FakeMarkdownPostprocessor:
    def __init__(self) -> None:
        self.calls: list[tuple[str, Path, Path, ConversionOptions]] = []

    def __call__(
        self,
        markdown: str,
        source_path: Path,
        output_path: Path,
        options: ConversionOptions,
    ) -> str:
        self.calls.append((markdown, source_path, output_path, options))
        return "# shaped by copilot"


class MarkItDownServiceTests(unittest.TestCase):
    def test_supported_extensions_are_case_insensitive(self) -> None:
        service = MarkItDownService(converter_factory=lambda: FakeConverter())
        self.assertTrue(service.is_supported("report.PDF"))
        self.assertTrue(service.is_supported("report.Docx"))
        self.assertFalse(service.is_supported("report.txt"))

    def test_validate_source_path_rejects_unsupported_extension(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "notes.txt"
            source_path.write_text("plain text", encoding="utf-8")

            service = MarkItDownService(converter_factory=lambda: FakeConverter())
            is_valid, message = service.validate_source_path(source_path)

            self.assertFalse(is_valid)
            self.assertIn("Unsupported file type", message)

    def test_build_output_path_appends_numeric_suffix(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "report.pdf"
            output_dir = temp_root / "output"
            source_path.write_bytes(b"%PDF-1.4")
            output_dir.mkdir()
            (output_dir / "report.md").write_text("existing", encoding="utf-8")
            (output_dir / "report (1).md").write_text("existing", encoding="utf-8")

            service = MarkItDownService(converter_factory=lambda: FakeConverter())
            next_path = service.build_output_path(source_path, output_dir, overwrite=False)

            self.assertEqual(next_path.name, "report (2).md")

    def test_convert_file_writes_markdown_output(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.pdf"
            source_path.write_bytes(b"%PDF-1.4")
            output_dir = temp_root / "markdown"
            converter = FakeConverter()
            service = MarkItDownService(converter_factory=lambda: converter)

            result = service.convert_file(
                source_path,
                ConversionOptions(output_dir=output_dir, keep_data_uris=True),
            )

            self.assertEqual(result.output_path.read_text(encoding="utf-8"), "# converted")
            self.assertEqual(result.title, "Sample")
            self.assertEqual(
                converter.calls,
                [(str(source_path.resolve()), {"keep_data_uris": True})],
            )

    def test_convert_file_wraps_missing_modules(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.pdf"
            source_path.write_bytes(b"%PDF-1.4")
            output_dir = temp_root / "markdown"
            service = MarkItDownService(converter_factory=lambda: MissingModuleConverter())

            with self.assertRaises(RuntimeDependencyError):
                service.convert_file(source_path, ConversionOptions(output_dir=output_dir))

    def test_convert_file_routes_xlsx_through_pdf_export(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.xlsx"
            source_path.write_bytes(b"PK\x03\x04")
            output_dir = temp_root / "markdown"
            markitdown_converter = FakeConverter()
            xlsx_pdf_exporter = FakeXlsxPdfExporter()
            service = MarkItDownService(
                converter_factory=lambda: markitdown_converter,
                xlsx_pdf_exporter=xlsx_pdf_exporter,
            )

            result = service.convert_file(
                source_path,
                ConversionOptions(output_dir=output_dir),
            )

            self.assertEqual(result.output_path.read_text(encoding="utf-8"), "# converted")
            self.assertEqual(len(markitdown_converter.calls), 1)
            converted_source, converted_options = markitdown_converter.calls[0]
            self.assertTrue(converted_source.endswith("sample.pdf"))
            self.assertEqual(converted_options, {"keep_data_uris": False})
            self.assertEqual(len(xlsx_pdf_exporter.calls), 1)
            exported_source, exported_pdf = xlsx_pdf_exporter.calls[0]
            self.assertEqual(exported_source, source_path.resolve())
            self.assertTrue(exported_pdf.name.endswith(".pdf"))

    def test_convert_file_propagates_xlsx_pdf_export_failures(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.xlsx"
            source_path.write_bytes(b"PK\x03\x04")
            output_dir = temp_root / "markdown"
            service = MarkItDownService(
                converter_factory=lambda: FakeConverter(),
                xlsx_pdf_exporter=MissingXlsxPdfExporter(),
            )

            with self.assertRaises(XlsxPdfExportError):
                service.convert_file(source_path, ConversionOptions(output_dir=output_dir))

    def test_convert_file_applies_markdown_postprocessor(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.pdf"
            source_path.write_bytes(b"%PDF-1.4")
            output_dir = temp_root / "markdown"
            converter = FakeConverter()
            postprocessor = FakeMarkdownPostprocessor()
            service = MarkItDownService(
                converter_factory=lambda: converter,
                markdown_postprocessor=postprocessor,
            )

            result = service.convert_file(
                source_path,
                ConversionOptions(
                    output_dir=output_dir,
                    copilot_formatting=True,
                    copilot_command="copilot",
                ),
            )

            self.assertEqual(result.markdown, "# shaped by copilot")
            self.assertEqual(
                result.output_path.read_text(encoding="utf-8"),
                "# shaped by copilot",
            )
            self.assertEqual(len(postprocessor.calls), 1)
            self.assertEqual(postprocessor.calls[0][0], "# converted")
            self.assertEqual(postprocessor.calls[0][1], source_path.resolve())
            self.assertEqual(postprocessor.calls[0][2], output_dir.resolve() / "sample.md")
            self.assertTrue(postprocessor.calls[0][3].copilot_formatting)

    def test_convert_file_reports_progress_stages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_path = temp_root / "sample.xlsx"
            source_path.write_bytes(b"PK\x03\x04")
            output_dir = temp_root / "markdown"
            service = MarkItDownService(
                converter_factory=lambda: FakeConverter(),
                xlsx_pdf_exporter=FakeXlsxPdfExporter(),
                markdown_postprocessor=FakeMarkdownPostprocessor(),
            )
            progress_events: list[tuple[str, float]] = []

            service.convert_file(
                source_path,
                ConversionOptions(
                    output_dir=output_dir,
                    copilot_formatting=True,
                ),
                progress_callback=lambda stage, value: progress_events.append((stage, value)),
            )

            self.assertEqual(
                [stage for stage, _ in progress_events],
                [
                    "validating",
                    "xlsx_pdf",
                    "markitdown",
                    "copilot",
                    "writing",
                    "finalizing",
                ],
            )
            self.assertEqual(progress_events[-1], ("finalizing", 1.0))


if __name__ == "__main__":
    unittest.main()
