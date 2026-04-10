from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Protocol

from ._copilot_formatter import CopilotCliFormatter
from ._temp_cleanup import XLSX_PDF_TEMP_PREFIX, build_temp_dir_prefix
from ._xlsx_pdf_exporter import default_xlsx_pdf_exporter

SUPPORTED_FILE_TYPES: dict[str, str] = {
    ".pdf": "PDF",
    ".docx": "DOCX",
    ".xlsx": "XLSX",
    ".html": "HTML",
    ".htm": "HTML",
}


class ConverterLike(Protocol):
    def convert(self, source: str, **kwargs: Any) -> Any:
        ...


ConverterFactory = Callable[[], ConverterLike]
XlsxPdfExporter = Callable[[Path, Path], None]
ProgressCallback = Callable[[str, float], None]


class RuntimeDependencyError(RuntimeError):
    pass


@dataclass(frozen=True)
class ConversionOptions:
    output_dir: Path
    overwrite: bool = False
    keep_data_uris: bool = False
    copilot_formatting: bool = False
    copilot_command: str = ""


@dataclass(frozen=True)
class ConversionResult:
    source_path: Path
    output_path: Path
    markdown: str
    title: str | None = None


MarkdownPostProcessor = Callable[[str, Path, Path, ConversionOptions], str]


def default_markdown_postprocessor(
    markdown: str,
    source_path: Path,
    output_path: Path,
    options: ConversionOptions,
) -> str:
    if not options.copilot_formatting:
        return markdown

    formatter = CopilotCliFormatter(command=options.copilot_command or None)
    return formatter.format_markdown(
        markdown,
        source_path=source_path,
        output_path=output_path,
    )


def default_converter_factory() -> ConverterLike:
    try:
        from markitdown import MarkItDown
    except ImportError as exc:
        raise RuntimeDependencyError(
            "MarkItDown is not installed. Install markitdown[pdf,docx,xlsx] before running the GUI."
        ) from exc

    return MarkItDown(enable_plugins=False)


class MarkItDownService:
    def __init__(
        self,
        converter_factory: ConverterFactory = default_converter_factory,
        xlsx_pdf_exporter: XlsxPdfExporter = default_xlsx_pdf_exporter,
        markdown_postprocessor: MarkdownPostProcessor = default_markdown_postprocessor,
    ):
        self._converter_factory = converter_factory
        self._xlsx_pdf_exporter = xlsx_pdf_exporter
        self._markdown_postprocessor = markdown_postprocessor
        self._converter: ConverterLike | None = None

    @property
    def supported_extensions(self) -> tuple[str, ...]:
        return tuple(SUPPORTED_FILE_TYPES.keys())

    def file_type_label(self, source_path: str | Path) -> str:
        return SUPPORTED_FILE_TYPES.get(Path(source_path).suffix.lower(), "Unsupported")

    def is_supported(self, source_path: str | Path) -> bool:
        return Path(source_path).suffix.lower() in SUPPORTED_FILE_TYPES

    def validate_source_path(self, source_path: str | Path) -> tuple[bool, str]:
        path = Path(source_path).expanduser()

        if not path.exists():
            return False, "File does not exist."
        if not path.is_file():
            return False, "Folders are not supported."
        if not self.is_supported(path):
            suffix = path.suffix.lower() or "[no extension]"
            return False, f"Unsupported file type: {suffix}"

        return True, self.file_type_label(path)

    def build_output_path(
        self,
        source_path: str | Path,
        output_dir: str | Path,
        *,
        overwrite: bool,
    ) -> Path:
        source = Path(source_path)
        target_dir = Path(output_dir)
        stem = source.stem or source.name or "converted"
        candidate = target_dir / f"{stem}.md"

        if overwrite:
            return candidate

        counter = 1
        while candidate.exists():
            candidate = target_dir / f"{stem} ({counter}).md"
            counter += 1

        return candidate

    def convert_file(
        self,
        source_path: str | Path,
        options: ConversionOptions,
        progress_callback: ProgressCallback | None = None,
    ) -> ConversionResult:
        source = Path(source_path).expanduser().resolve()
        self._report_progress(progress_callback, "validating", 0.05)
        is_valid, validation_message = self.validate_source_path(source)
        if not is_valid:
            raise ValueError(validation_message)

        output_dir = options.output_dir.expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = self.build_output_path(
            source,
            output_dir,
            overwrite=options.overwrite,
        )

        if source.suffix.lower() == ".xlsx":
            markdown, title = self._convert_xlsx_via_pdf(
                source,
                options,
                progress_callback,
            )

            if options.copilot_formatting:
                self._report_progress(progress_callback, "copilot", 0.88)

            markdown = self._markdown_postprocessor(
                markdown,
                source,
                output_path,
                options,
            )
            self._report_progress(progress_callback, "writing", 0.96)
            output_path.write_text(markdown, encoding="utf-8")
            self._report_progress(progress_callback, "finalizing", 1.0)

            return ConversionResult(
                source_path=source,
                output_path=output_path,
                markdown=markdown,
                title=title,
            )

        self._report_progress(
            progress_callback,
            "markitdown",
            0.55 if options.copilot_formatting else 0.78,
        )
        markdown, title = self._convert_with_markitdown(source, options)

        if options.copilot_formatting:
            self._report_progress(progress_callback, "copilot", 0.88)
        markdown = self._markdown_postprocessor(
            markdown,
            source,
            output_path,
            options,
        )
        self._report_progress(progress_callback, "writing", 0.96)
        output_path.write_text(markdown, encoding="utf-8")
        self._report_progress(progress_callback, "finalizing", 1.0)

        return ConversionResult(
            source_path=source,
            output_path=output_path,
            markdown=markdown,
            title=title,
        )

    def path_key(self, source_path: str | Path) -> str:
        path = Path(source_path).expanduser()
        return os.path.normcase(str(path.resolve(strict=False)))

    def _convert_xlsx_via_pdf(
        self,
        source: Path,
        options: ConversionOptions,
        progress_callback: ProgressCallback | None = None,
    ) -> tuple[str, str | None]:
        with tempfile.TemporaryDirectory(
            prefix=build_temp_dir_prefix(XLSX_PDF_TEMP_PREFIX)
        ) as temp_dir:
            pdf_path = Path(temp_dir) / f"{source.stem}.pdf"
            self._report_progress(progress_callback, "xlsx_pdf", 0.30)
            self._xlsx_pdf_exporter(source, pdf_path)
            self._report_progress(
                progress_callback,
                "markitdown",
                0.65 if options.copilot_formatting else 0.86,
            )
            return self._convert_with_markitdown(pdf_path, options)

    def _convert_with_markitdown(
        self,
        source: Path,
        options: ConversionOptions,
    ) -> tuple[str, str | None]:
        converter = self._get_converter()
        try:
            result = converter.convert(
                str(source),
                keep_data_uris=options.keep_data_uris,
            )
        except ModuleNotFoundError as exc:
            raise RuntimeDependencyError(
                "A required optional dependency is missing. Install markitdown[pdf,docx,xlsx]."
            ) from exc
        except ImportError as exc:
            raise RuntimeDependencyError(
                "A runtime dependency could not be imported. Reinstall this package and markitdown[pdf,docx,xlsx]."
            ) from exc

        markdown = getattr(result, "markdown", None)
        if markdown is None:
            markdown = getattr(result, "text_content", None)

        if not isinstance(markdown, str):
            raise RuntimeError("MarkItDown returned an unexpected result object.")

        title = getattr(result, "title", None)
        if title is not None and not isinstance(title, str):
            title = str(title)

        return markdown, title

    def _get_converter(self) -> ConverterLike:
        if self._converter is None:
            self._converter = self._converter_factory()
        return self._converter

    def _report_progress(
        self,
        progress_callback: ProgressCallback | None,
        stage_key: str,
        file_progress: float,
    ) -> None:
        if progress_callback is None:
            return
        bounded_progress = min(max(file_progress, 0.0), 1.0)
        progress_callback(stage_key, bounded_progress)
