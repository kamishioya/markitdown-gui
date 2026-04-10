from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from ._service import ConversionOptions, MarkItDownService


class ConversionWorker(QObject):
    file_started = Signal(str)
    file_succeeded = Signal(str, str, str)
    file_failed = Signal(str, str)
    progress_changed = Signal(int, int)
    stage_changed = Signal(str, int, int, int, str)
    log_message = Signal(str)
    finished = Signal(int, int, bool)

    def __init__(
        self,
        source_paths: list[Path],
        options: ConversionOptions,
        service: MarkItDownService,
    ) -> None:
        super().__init__()
        self._source_paths = source_paths
        self._options = options
        self._service = service
        self._cancel_requested = False

    @Slot()
    def run(self) -> None:
        succeeded = 0
        failed = 0
        total = len(self._source_paths)

        for index, source_path in enumerate(self._source_paths, start=1):
            if self._cancel_requested:
                self.log_message.emit("cancel_waiting")
                break

            self.file_started.emit(str(source_path))
            progress_callback = self._make_progress_callback(source_path, index, total)
            progress_callback("starting", 0.0)
            try:
                result = self._service.convert_file(
                    source_path,
                    self._options,
                    progress_callback=progress_callback,
                )
            except Exception as exc:
                failed += 1
                progress_callback("failed", 1.0)
                self.file_failed.emit(str(source_path), str(exc))
            else:
                succeeded += 1
                self.file_succeeded.emit(
                    str(result.source_path),
                    str(result.output_path),
                    result.markdown,
                )

            self.progress_changed.emit(index, total)

        self.finished.emit(succeeded, failed, self._cancel_requested)

    @Slot()
    def cancel(self) -> None:
        self._cancel_requested = True

    def _make_progress_callback(
        self,
        source_path: Path,
        current_index: int,
        total_files: int,
    ):
        def callback(stage_key: str, file_progress: float) -> None:
            bounded_progress = min(max(file_progress, 0.0), 1.0)
            if total_files <= 0:
                overall_percent = 100
            else:
                overall_percent = int(
                    round(((current_index - 1) + bounded_progress) / total_files * 100)
                )
            overall_percent = min(max(overall_percent, 0), 100)
            self.stage_changed.emit(
                str(source_path),
                current_index,
                total_files,
                overall_percent,
                stage_key,
            )

        return callback
