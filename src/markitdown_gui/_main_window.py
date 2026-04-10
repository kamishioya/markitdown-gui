from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, QSignalBlocker, QThread, Qt, QUrl
from PySide6.QtGui import QAction, QColor, QDesktopServices, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QPlainTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ._copilot_formatter import (
    COPILOT_AUTH_MESSAGE,
    COPILOT_EMPTY_OUTPUT_MESSAGE,
    COPILOT_FAILURE_PREFIX,
    COPILOT_MISSING_MESSAGE,
    COPILOT_TIMEOUT_MESSAGE,
    detect_copilot_cli_command,
)
from ._copilot_setup_dialog import CopilotSetupDialog
from ._localization import DEFAULT_LANGUAGE, LANGUAGE_NAMES, get_text
from ._service import ConversionOptions, MarkItDownService
from ._worker import ConversionWorker
from ._xlsx_pdf_exporter import (
    XLSX_PDF_EXCEL_MISSING_MESSAGE,
    XLSX_PDF_FAILURE_PREFIX,
    XLSX_PDF_OUTPUT_MISSING_MESSAGE,
    XLSX_PDF_POWERSHELL_MISSING_MESSAGE,
    XLSX_PDF_TIMEOUT_MESSAGE,
)

MARKDOWN_ROLE = Qt.ItemDataRole.UserRole + 1
ERROR_ROLE = Qt.ItemDataRole.UserRole + 2
SOURCE_ROLE = Qt.ItemDataRole.UserRole + 3
STATUS_ROLE = Qt.ItemDataRole.UserRole + 4

SETTINGS_ORGANIZATION = "MarkItDown"
SETTINGS_APPLICATION = "MarkItDown GUI"
LANGUAGE_SETTING_KEY = "ui/language"
COPILOT_ENABLED_SETTING_KEY = "copilot/enabled"
COPILOT_COMMAND_SETTING_KEY = "copilot/command"

STATUS_READY = "ready"
STATUS_PROCESSING = "processing"
STATUS_DONE = "done"
STATUS_FAILED = "failed"


class MainWindow(QMainWindow):
    def __init__(
        self,
        settings: QSettings | None = None,
        default_copilot_command: str | None = None,
    ) -> None:
        super().__init__()
        self._service = MarkItDownService()
        self._thread: QThread | None = None
        self._worker: ConversionWorker | None = None
        self._row_by_source_key: dict[str, int] = {}
        self._default_copilot_command = default_copilot_command
        self._busy_state = False
        self._current_progress_percent = 0
        self._current_progress_stage_key = ""
        self._current_progress_source_path = ""
        self._current_progress_current = 0
        self._current_progress_total = 0
        self._settings = settings or QSettings(
            SETTINGS_ORGANIZATION,
            SETTINGS_APPLICATION,
        )
        self._language = self._load_language()

        self.resize(1100, 760)
        self.setAcceptDrops(True)

        self._build_ui()
        self._load_copilot_settings()
        self._apply_language()
        self._set_default_output_dir()
        self._update_actions()
        self._update_status_bar()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        local_files = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.isLocalFile()
        ]
        if local_files:
            self._add_paths(local_files)
            event.acceptProposedAction()
            return
        event.ignore()

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self._thread is not None and self._thread.isRunning():
            self._append_log(self._text("log_close_requested"))
            self._request_cancel()
            event.ignore()
            return
        super().closeEvent(event)

    def _build_ui(self) -> None:
        menu_bar = self.menuBar()
        menu_bar.setNativeMenuBar(False)
        self._settings_menu = menu_bar.addMenu("")
        self._copilot_setup_action = QAction(self)
        self._copilot_setup_action.triggered.connect(self._open_copilot_setup_dialog)
        self._settings_menu.addAction(self._copilot_setup_action)

        central_widget = QWidget(self)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        self._intro_label = QLabel(self)
        self._intro_label.setWordWrap(True)
        main_layout.addWidget(self._intro_label)

        settings_row = QHBoxLayout()
        settings_row.addStretch(1)
        self._language_label = QLabel(self)
        self._language_combo = QComboBox(self)
        for language_code, language_name in LANGUAGE_NAMES.items():
            self._language_combo.addItem(language_name, language_code)
        self._set_language_combo_value(self._language)
        self._language_combo.currentIndexChanged.connect(self._on_language_changed)
        settings_row.addWidget(self._language_label)
        settings_row.addWidget(self._language_combo)
        main_layout.addLayout(settings_row)

        self._file_group = QGroupBox(self)
        file_layout = QVBoxLayout(self._file_group)
        file_button_layout = QHBoxLayout()
        self._add_files_button = QPushButton(self)
        self._remove_files_button = QPushButton(self)
        self._clear_files_button = QPushButton(self)
        file_button_layout.addWidget(self._add_files_button)
        file_button_layout.addWidget(self._remove_files_button)
        file_button_layout.addWidget(self._clear_files_button)
        file_button_layout.addStretch(1)
        file_layout.addLayout(file_button_layout)

        self._file_table = QTableWidget(0, 4, self)
        self._file_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._file_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._file_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._file_table.setAlternatingRowColors(True)
        self._file_table.verticalHeader().setVisible(False)
        header = self._file_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.resizeSection(0, 420)
        header.resizeSection(1, 90)
        header.resizeSection(2, 110)
        file_layout.addWidget(self._file_table)
        main_layout.addWidget(self._file_group, stretch=3)

        self._output_group = QGroupBox(self)
        output_layout = QVBoxLayout(self._output_group)
        output_row = QHBoxLayout()
        self._output_dir_edit = QLineEdit(self)
        self._browse_output_button = QPushButton(self)
        self._output_dir_label = QLabel(self)
        output_row.addWidget(self._output_dir_label)
        output_row.addWidget(self._output_dir_edit, stretch=1)
        output_row.addWidget(self._browse_output_button)
        output_layout.addLayout(output_row)

        option_row = QHBoxLayout()
        self._overwrite_checkbox = QCheckBox(self)
        self._keep_data_uris_checkbox = QCheckBox(self)
        self._open_output_checkbox = QCheckBox(self)
        option_row.addWidget(self._overwrite_checkbox)
        option_row.addWidget(self._keep_data_uris_checkbox)
        option_row.addWidget(self._open_output_checkbox)
        option_row.addStretch(1)
        output_layout.addLayout(option_row)

        self._copilot_checkbox = QCheckBox(self)
        output_layout.addWidget(self._copilot_checkbox)

        copilot_row = QHBoxLayout()
        self._copilot_command_label = QLabel(self)
        self._copilot_command_edit = QLineEdit(self)
        self._browse_copilot_button = QPushButton(self)
        copilot_row.addWidget(self._copilot_command_label)
        copilot_row.addWidget(self._copilot_command_edit, stretch=1)
        copilot_row.addWidget(self._browse_copilot_button)
        output_layout.addLayout(copilot_row)
        main_layout.addWidget(self._output_group)

        action_row = QHBoxLayout()
        self._convert_button = QPushButton(self)
        self._cancel_button = QPushButton(self)
        self._cancel_button.setEnabled(False)
        action_row.addStretch(1)
        action_row.addWidget(self._convert_button)
        action_row.addWidget(self._cancel_button)
        main_layout.addLayout(action_row)

        self._tabs = QTabWidget(self)
        self._preview_edit = QPlainTextEdit(self)
        self._preview_edit.setReadOnly(True)
        self._log_edit = QPlainTextEdit(self)
        self._log_edit.setReadOnly(True)
        self._tabs.addTab(self._preview_edit, "")
        self._tabs.addTab(self._log_edit, "")
        main_layout.addWidget(self._tabs, stretch=2)

        self.setCentralWidget(central_widget)

        self._status_progress_bar = QProgressBar(self)
        self._status_progress_bar.setRange(0, 100)
        self._status_progress_bar.setValue(0)
        self._status_progress_bar.setFormat("%p%")
        self._status_progress_bar.setFixedWidth(160)
        self.statusBar().addPermanentWidget(self._status_progress_bar)
        self._status_progress_bar.hide()

        self._add_files_button.clicked.connect(self._pick_files)
        self._remove_files_button.clicked.connect(self._remove_selected_rows)
        self._clear_files_button.clicked.connect(self._clear_rows)
        self._browse_output_button.clicked.connect(self._pick_output_dir)
        self._browse_copilot_button.clicked.connect(self._pick_copilot_command)
        self._convert_button.clicked.connect(self._start_conversion)
        self._cancel_button.clicked.connect(self._request_cancel)
        self._file_table.itemSelectionChanged.connect(self._update_preview)
        self._output_dir_edit.textChanged.connect(self._update_actions)
        self._copilot_checkbox.toggled.connect(self._on_copilot_toggled)
        self._copilot_command_edit.textChanged.connect(self._on_copilot_command_changed)

    def _pick_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            self._text("dialog_select_files"),
            str(Path.home()),
            self._text("dialog_file_filter"),
        )
        if file_paths:
            self._add_paths(file_paths)

    def _pick_output_dir(self) -> None:
        current = self._output_dir_edit.text().strip() or str(Path.home())
        output_dir = QFileDialog.getExistingDirectory(
            self,
            self._text("dialog_select_output"),
            current,
        )
        if output_dir:
            self._output_dir_edit.setText(output_dir)

    def _pick_copilot_command(self) -> None:
        current = self._copilot_command_edit.text().strip() or str(Path.home())
        command_path, _ = QFileDialog.getOpenFileName(
            self,
            self._text("dialog_select_copilot_command"),
            current,
            self._text("dialog_copilot_command_filter"),
        )
        if command_path:
            self._copilot_command_edit.setText(command_path)

    def _open_copilot_setup_dialog(self) -> None:
        dialog = CopilotSetupDialog(
            locale=self._language,
            copilot_enabled=self._copilot_checkbox.isChecked(),
            copilot_command=self._copilot_command_edit.text().strip(),
            default_copilot_command=self._resolve_default_copilot_command(),
            parent=self,
        )
        if dialog.exec() != dialog.DialogCode.Accepted:
            return

        self._copilot_checkbox.setChecked(dialog.copilot_enabled())
        self._copilot_command_edit.setText(dialog.copilot_command())
        self._append_log(self._text("log_copilot_settings_updated"))

    def _add_paths(self, raw_paths: list[str]) -> None:
        added_count = 0
        skipped_count = 0

        for raw_path in raw_paths:
            source_path = Path(raw_path)
            source_key = self._service.path_key(source_path)
            if source_key in self._row_by_source_key:
                skipped_count += 1
                self._append_log(
                    self._text("log_skipped_duplicate", path=source_path)
                )
                continue

            is_valid, validation_message = self._service.validate_source_path(source_path)
            if not is_valid:
                skipped_count += 1
                self._append_log(
                    self._text(
                        "log_skipped_path",
                        path=source_path,
                        reason=self._localize_message(validation_message),
                    )
                )
                continue

            row = self._file_table.rowCount()
            self._file_table.insertRow(row)

            source_item = QTableWidgetItem(str(source_path.resolve()))
            source_item.setData(SOURCE_ROLE, str(source_path.resolve()))
            source_item.setData(MARKDOWN_ROLE, "")
            source_item.setData(ERROR_ROLE, "")
            type_item = QTableWidgetItem(validation_message)
            status_item = QTableWidgetItem()
            output_item = QTableWidgetItem("")

            self._file_table.setItem(row, 0, source_item)
            self._file_table.setItem(row, 1, type_item)
            self._file_table.setItem(row, 2, status_item)
            self._file_table.setItem(row, 3, output_item)

            self._row_by_source_key[source_key] = row
            self._set_row_status(row, STATUS_READY)
            added_count += 1

        if added_count:
            self._append_log(self._text("log_added_count", count=added_count))
        if skipped_count:
            self._append_log(self._text("log_skipped_count", count=skipped_count))

        if self._file_table.rowCount() > 0 and self._file_table.currentRow() < 0:
            self._file_table.selectRow(0)

        self._update_actions()
        self._update_status_bar()

    def _remove_selected_rows(self) -> None:
        rows = sorted({item.row() for item in self._file_table.selectedItems()}, reverse=True)
        if not rows:
            return

        for row in rows:
            self._remove_row(row)

        self._rebuild_row_map()
        self._update_actions()
        self._update_preview()
        self._update_status_bar()

    def _clear_rows(self) -> None:
        self._file_table.setRowCount(0)
        self._row_by_source_key.clear()
        self._preview_edit.clear()
        self._append_log(self._text("log_cleared"))
        self._update_actions()
        self._update_status_bar()

    def _remove_row(self, row: int) -> None:
        source_item = self._file_table.item(row, 0)
        if source_item is not None:
            source_value = source_item.data(SOURCE_ROLE)
            if isinstance(source_value, str):
                self._row_by_source_key.pop(self._service.path_key(source_value), None)
        self._file_table.removeRow(row)

    def _rebuild_row_map(self) -> None:
        self._row_by_source_key.clear()
        for row in range(self._file_table.rowCount()):
            source_item = self._file_table.item(row, 0)
            if source_item is None:
                continue
            source_value = source_item.data(SOURCE_ROLE)
            if isinstance(source_value, str):
                self._row_by_source_key[self._service.path_key(source_value)] = row

    def _set_default_output_dir(self) -> None:
        if self._output_dir_edit.text().strip():
            return
        self._output_dir_edit.setText(str(self._default_output_dir()))

    def _load_copilot_settings(self) -> None:
        checkbox_blocker = QSignalBlocker(self._copilot_checkbox)
        command_blocker = QSignalBlocker(self._copilot_command_edit)

        copilot_enabled = self._settings.value(
            COPILOT_ENABLED_SETTING_KEY,
            False,
            type=bool,
        )
        copilot_command = self._settings.value(
            COPILOT_COMMAND_SETTING_KEY,
            "",
            type=str,
        )

        self._copilot_checkbox.setChecked(copilot_enabled)
        self._copilot_command_edit.setText(
            copilot_command or self._resolve_default_copilot_command()
        )

        del command_blocker
        del checkbox_blocker
        self._update_copilot_controls()

    def _resolve_default_copilot_command(self) -> str:
        if self._default_copilot_command is not None:
            return self._default_copilot_command
        detected_command = detect_copilot_cli_command()
        return detected_command or ""

    def _gather_source_paths(self) -> list[Path]:
        paths: list[Path] = []
        for row in range(self._file_table.rowCount()):
            source_item = self._file_table.item(row, 0)
            if source_item is None:
                continue
            source_value = source_item.data(SOURCE_ROLE)
            if isinstance(source_value, str):
                paths.append(Path(source_value))
        return paths

    def _start_conversion(self) -> None:
        if self._thread is not None and self._thread.isRunning():
            return

        output_dir_text = self._output_dir_edit.text().strip()
        if not output_dir_text:
            QMessageBox.warning(
                self,
                self._text("warning_output_folder_required_title"),
                self._text("warning_output_folder_required_message"),
            )
            return

        source_paths = self._gather_source_paths()
        if not source_paths:
            return

        options = ConversionOptions(
            output_dir=Path(output_dir_text),
            overwrite=self._overwrite_checkbox.isChecked(),
            keep_data_uris=self._keep_data_uris_checkbox.isChecked(),
            copilot_formatting=self._copilot_checkbox.isChecked(),
            copilot_command=self._copilot_command_edit.text().strip(),
        )

        self._thread = QThread(self)
        self._worker = ConversionWorker(source_paths, options, MarkItDownService())
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.file_started.connect(self._on_file_started)
        self._worker.file_succeeded.connect(self._on_file_succeeded)
        self._worker.file_failed.connect(self._on_file_failed)
        self._worker.progress_changed.connect(self._on_progress_changed)
        self._worker.stage_changed.connect(self._on_stage_changed)
        self._worker.log_message.connect(self._on_worker_log_message)
        self._worker.finished.connect(self._on_conversion_finished)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._on_thread_cleaned_up)

        self._append_log(self._text("log_start_conversion", count=len(source_paths)))
        self._set_busy_state(True)
        self._thread.start()

    def _request_cancel(self) -> None:
        if self._worker is None:
            return
        self._cancel_button.setEnabled(False)
        self._worker.cancel()
        self._append_log(self._text("log_cancel_requested"))

    def _on_file_started(self, source_path: str) -> None:
        row = self._row_from_source(source_path)
        if row is not None:
            self._set_row_status(row, STATUS_PROCESSING)
        self._current_progress_source_path = source_path

    def _on_file_succeeded(self, source_path: str, output_path: str, markdown: str) -> None:
        row = self._row_from_source(source_path)
        if row is None:
            return

        source_item = self._file_table.item(row, 0)
        output_item = self._file_table.item(row, 3)
        if source_item is not None:
            source_item.setData(MARKDOWN_ROLE, markdown)
            source_item.setData(ERROR_ROLE, "")
        if output_item is not None:
            output_item.setText(output_path)

        self._set_row_status(row, STATUS_DONE)
        self._append_log(
            self._text("log_converted", source=source_path, output=output_path)
        )
        self._update_preview()
        self._update_status_bar()

    def _on_file_failed(self, source_path: str, error_message: str) -> None:
        row = self._row_from_source(source_path)
        if row is None:
            return

        localized_error_message = self._localize_message(error_message)

        source_item = self._file_table.item(row, 0)
        if source_item is not None:
            source_item.setData(MARKDOWN_ROLE, "")
            source_item.setData(ERROR_ROLE, localized_error_message)

        self._set_row_status(row, STATUS_FAILED)
        self._append_log(
            self._text(
                "log_failed",
                source=source_path,
                message=localized_error_message,
            )
        )
        self._update_preview()
        self._update_status_bar()

    def _on_progress_changed(self, current: int, total: int) -> None:
        self._current_progress_percent = max(self._current_progress_percent, int(round(current / total * 100)))
        self._current_progress_current = current
        self._current_progress_total = total
        self._update_status_bar()

    def _on_stage_changed(
        self,
        source_path: str,
        current: int,
        total: int,
        percent: int,
        stage_key: str,
    ) -> None:
        self._current_progress_source_path = source_path
        self._current_progress_current = current
        self._current_progress_total = total
        self._current_progress_percent = percent
        self._current_progress_stage_key = stage_key

        row = self._row_from_source(source_path)
        if row is not None:
            self._set_row_processing_progress(
                row,
                stage_key,
                percent,
            )

        self._update_status_bar()

    def _on_conversion_finished(self, succeeded: int, failed: int, cancelled: bool) -> None:
        if cancelled:
            self._append_log(self._text("log_batch_cancelled"))
        else:
            self._append_log(self._text("log_batch_finished"))

        self._append_log(
            self._text("log_summary", succeeded=succeeded, failed=failed)
        )

        if not cancelled and succeeded > 0 and self._open_output_checkbox.isChecked():
            output_dir = self._output_dir_edit.text().strip()
            if output_dir:
                QDesktopServices.openUrl(QUrl.fromLocalFile(output_dir))

        self._set_busy_state(False)
        self._update_status_bar()

    def _on_thread_cleaned_up(self) -> None:
        self._thread = None
        self._worker = None
        self._update_status_bar()

    def _on_worker_log_message(self, message: str) -> None:
        if message == "cancel_waiting":
            self._append_log(self._text("log_cancellation_wait"))
            return
        self._append_log(message)

    def _on_language_changed(self, index: int) -> None:
        language_code = self._language_combo.itemData(index)
        if not isinstance(language_code, str) or language_code == self._language:
            return

        old_default_output_dir = self._default_output_dir(self._language)
        should_update_default_output_dir = (
            self._output_dir_edit.text().strip() == str(old_default_output_dir)
        )

        self._language = language_code
        self._settings.setValue(LANGUAGE_SETTING_KEY, language_code)
        self._apply_language()

        if should_update_default_output_dir:
            self._output_dir_edit.setText(str(self._default_output_dir()))

        self._append_log(
            self._text(
                "log_language_changed",
                language=LANGUAGE_NAMES.get(language_code, language_code),
            )
        )

    def _on_copilot_toggled(self, checked: bool) -> None:
        self._settings.setValue(COPILOT_ENABLED_SETTING_KEY, checked)
        self._update_copilot_controls()

    def _on_copilot_command_changed(self, command: str) -> None:
        self._settings.setValue(COPILOT_COMMAND_SETTING_KEY, command.strip())

    def _row_from_source(self, source_path: str) -> int | None:
        return self._row_by_source_key.get(self._service.path_key(source_path))

    def _set_row_status(self, row: int, status: str) -> None:
        status_item = self._file_table.item(row, 2)
        if status_item is None:
            return

        status_item.setData(STATUS_ROLE, status)
        status_item.setText(self._status_text(status))
        color = QColor("#4b5563")
        if status == STATUS_DONE:
            color = QColor("#0f766e")
        elif status == STATUS_FAILED:
            color = QColor("#b91c1c")
        elif status == STATUS_PROCESSING:
            color = QColor("#9a3412")
        status_item.setForeground(color)

    def _set_row_processing_progress(self, row: int, stage_key: str, overall_percent: int) -> None:
        status_item = self._file_table.item(row, 2)
        if status_item is None:
            return

        status_item.setData(STATUS_ROLE, STATUS_PROCESSING)
        status_item.setText(
            self._text(
                "status_processing_detail_short",
                stage=self._progress_stage_text(stage_key),
                percent=overall_percent,
            )
        )
        status_item.setForeground(QColor("#9a3412"))

    def _set_busy_state(self, is_busy: bool) -> None:
        self._busy_state = is_busy
        self._add_files_button.setEnabled(not is_busy)
        self._remove_files_button.setEnabled(not is_busy and self._file_table.rowCount() > 0)
        self._clear_files_button.setEnabled(not is_busy and self._file_table.rowCount() > 0)
        self._browse_output_button.setEnabled(not is_busy)
        self._output_dir_edit.setEnabled(not is_busy)
        self._language_combo.setEnabled(not is_busy)
        self._overwrite_checkbox.setEnabled(not is_busy)
        self._keep_data_uris_checkbox.setEnabled(not is_busy)
        self._open_output_checkbox.setEnabled(not is_busy)
        self._copilot_checkbox.setEnabled(not is_busy)
        self._convert_button.setEnabled(not is_busy and self._file_table.rowCount() > 0)
        self._cancel_button.setEnabled(is_busy)
        self._update_copilot_controls()
        if is_busy:
            self._status_progress_bar.show()
        else:
            self._status_progress_bar.hide()
            self._status_progress_bar.setValue(0)
            self._current_progress_percent = 0
            self._current_progress_stage_key = ""
            self._current_progress_source_path = ""
            self._current_progress_current = 0
            self._current_progress_total = 0

    def _update_actions(self) -> None:
        has_rows = self._file_table.rowCount() > 0
        has_output_dir = bool(self._output_dir_edit.text().strip())
        is_busy = self._is_busy()

        self._add_files_button.setEnabled(not is_busy)
        self._remove_files_button.setEnabled(not is_busy and has_rows)
        self._clear_files_button.setEnabled(not is_busy and has_rows)
        self._browse_output_button.setEnabled(not is_busy)
        self._language_combo.setEnabled(not is_busy)
        self._copilot_checkbox.setEnabled(not is_busy)
        self._convert_button.setEnabled(not is_busy and has_rows and has_output_dir)
        self._cancel_button.setEnabled(is_busy)
        self._update_copilot_controls()

    def _update_copilot_controls(self) -> None:
        is_busy = self._is_busy()
        copilot_inputs_enabled = self._copilot_checkbox.isChecked() and not is_busy
        self._copilot_command_label.setEnabled(copilot_inputs_enabled)
        self._copilot_command_edit.setEnabled(copilot_inputs_enabled)
        self._browse_copilot_button.setEnabled(copilot_inputs_enabled)

    def _update_preview(self) -> None:
        selected_rows = {item.row() for item in self._file_table.selectedItems()}
        if not selected_rows:
            self._preview_edit.setPlainText(self._text("preview_select_row"))
            return

        row = min(selected_rows)
        source_item = self._file_table.item(row, 0)
        if source_item is None:
            self._preview_edit.clear()
            return

        markdown = source_item.data(MARKDOWN_ROLE)
        error_message = source_item.data(ERROR_ROLE)
        if isinstance(markdown, str) and markdown:
            self._preview_edit.setPlainText(markdown)
            return
        if isinstance(error_message, str) and error_message:
            self._preview_edit.setPlainText(
                f"{self._text('preview_failed_prefix')}\n\n{error_message}"
            )
            return

        self._preview_edit.setPlainText(self._text("preview_no_output"))

    def _append_log(self, message: str) -> None:
        self._log_edit.appendPlainText(message)
        self._tabs.setCurrentWidget(self._log_edit)

    def _update_status_bar(self) -> None:
        is_busy = self._is_busy()
        if is_busy and self._current_progress_total > 0:
            self._status_progress_bar.show()
            self._status_progress_bar.setValue(self._current_progress_percent)
            self.statusBar().showMessage(
                self._text(
                    "status_progress_detail",
                    percent=self._current_progress_percent,
                    current=self._current_progress_current,
                    total=self._current_progress_total,
                    stage=self._progress_stage_text(self._current_progress_stage_key),
                    name=Path(self._current_progress_source_path).name or self._current_progress_source_path,
                )
            )
            return

        ready = 0
        done = 0
        failed = 0
        processing = 0

        for row in range(self._file_table.rowCount()):
            status_item = self._file_table.item(row, 2)
            if status_item is None:
                continue
            status_code = self._status_code_from_item(status_item)
            if status_code == STATUS_READY:
                ready += 1
            elif status_code == STATUS_DONE:
                done += 1
            elif status_code == STATUS_FAILED:
                failed += 1
            elif status_code == STATUS_PROCESSING:
                processing += 1

        self.statusBar().showMessage(
            self._text(
                "status_bar_summary",
                total=self._file_table.rowCount(),
                ready=ready,
                processing=processing,
                done=done,
                failed=failed,
            )
        )

    def _progress_stage_text(self, stage_key: str) -> str:
        stage_text_keys = {
            "starting": "progress_stage_starting",
            "validating": "progress_stage_validating",
            "xlsx_pdf": "progress_stage_xlsx_pdf",
            "markitdown": "progress_stage_markitdown",
            "copilot": "progress_stage_copilot",
            "writing": "progress_stage_writing",
            "finalizing": "progress_stage_finalizing",
            "failed": "progress_stage_failed",
        }
        text_key = stage_text_keys.get(stage_key)
        if text_key is None:
            return self._text("status_processing")
        return self._text(text_key)

    def _is_busy(self) -> bool:
        return self._busy_state or (self._thread is not None and self._thread.isRunning())

    def _apply_language(self) -> None:
        self.setWindowTitle(self._text("window_title"))
        self._settings_menu.setTitle(self._text("menu_settings"))
        self._copilot_setup_action.setText(self._text("action_copilot_setup"))
        self._intro_label.setText(self._text("intro_text"))
        self._language_label.setText(self._text("language_label"))
        self._file_group.setTitle(self._text("input_group"))
        self._add_files_button.setText(self._text("button_add_files"))
        self._remove_files_button.setText(self._text("button_remove_selected"))
        self._clear_files_button.setText(self._text("button_clear_list"))
        self._file_table.setHorizontalHeaderLabels(
            [
                self._text("header_source"),
                self._text("header_type"),
                self._text("header_status"),
                self._text("header_output"),
            ]
        )
        self._output_group.setTitle(self._text("output_group"))
        self._output_dir_label.setText(self._text("output_folder_label"))
        self._browse_output_button.setText(self._text("button_browse"))
        self._overwrite_checkbox.setText(self._text("checkbox_overwrite"))
        self._keep_data_uris_checkbox.setText(self._text("checkbox_keep_data_uris"))
        self._open_output_checkbox.setText(self._text("checkbox_open_output"))
        self._copilot_checkbox.setText(self._text("checkbox_copilot_format"))
        self._copilot_command_label.setText(self._text("copilot_command_label"))
        self._copilot_command_edit.setPlaceholderText(
            self._text("copilot_command_placeholder")
        )
        self._browse_copilot_button.setText(self._text("button_browse"))
        self._convert_button.setText(self._text("button_convert"))
        self._cancel_button.setText(self._text("button_cancel"))
        self._tabs.setTabText(
            self._tabs.indexOf(self._preview_edit),
            self._text("tab_preview"),
        )
        self._tabs.setTabText(
            self._tabs.indexOf(self._log_edit),
            self._text("tab_log"),
        )
        self._refresh_status_labels()
        self._update_preview()
        self._update_status_bar()

    def _refresh_status_labels(self) -> None:
        for row in range(self._file_table.rowCount()):
            status_item = self._file_table.item(row, 2)
            if status_item is None:
                continue
            status_code = self._status_code_from_item(status_item)
            if status_code is None:
                continue
            status_item.setText(self._status_text(status_code))

    def _status_code_from_item(self, status_item: QTableWidgetItem) -> str | None:
        status_code = status_item.data(STATUS_ROLE)
        if isinstance(status_code, str):
            return status_code
        return None

    def _status_text(self, status: str) -> str:
        return self._text(f"status_{status}")

    def _load_language(self) -> str:
        language = self._settings.value(LANGUAGE_SETTING_KEY, DEFAULT_LANGUAGE, type=str)
        if language in LANGUAGE_NAMES:
            return language
        return DEFAULT_LANGUAGE

    def _set_language_combo_value(self, language: str) -> None:
        blocker = QSignalBlocker(self._language_combo)
        index = self._language_combo.findData(language)
        if index < 0:
            index = self._language_combo.findData(DEFAULT_LANGUAGE)
        if index >= 0:
            self._language_combo.setCurrentIndex(index)
        del blocker

    def _default_output_dir(self, language: str | None = None) -> Path:
        selected_language = language or self._language
        return Path.home() / "Documents" / get_text(
            selected_language,
            "default_output_folder_name",
        )

    def _localize_message(self, message: str) -> str:
        known_messages = {
            "File does not exist.": self._text("validation_file_missing"),
            "Folders are not supported.": self._text("validation_folder_not_supported"),
            "MarkItDown is not installed. Install markitdown[pdf,docx,xlsx] before running the GUI.": self._text("runtime_markitdown_missing"),
            "A required optional dependency is missing. Install markitdown[pdf,docx,xlsx].": self._text("runtime_optional_dependency_missing"),
            "A runtime dependency could not be imported. Reinstall this package and markitdown[pdf,docx,xlsx].": self._text("runtime_dependency_import_failed"),
            "MarkItDown returned an unexpected result object.": self._text("runtime_unexpected_result"),
            COPILOT_MISSING_MESSAGE: self._text("runtime_copilot_missing"),
            COPILOT_AUTH_MESSAGE: self._text("runtime_copilot_auth"),
            COPILOT_EMPTY_OUTPUT_MESSAGE: self._text("runtime_copilot_empty_output"),
            COPILOT_TIMEOUT_MESSAGE: self._text("runtime_copilot_timeout"),
            XLSX_PDF_POWERSHELL_MISSING_MESSAGE: self._text("runtime_xlsx_pdf_powershell_missing"),
            XLSX_PDF_EXCEL_MISSING_MESSAGE: self._text("runtime_xlsx_pdf_excel_missing"),
            XLSX_PDF_OUTPUT_MISSING_MESSAGE: self._text("runtime_xlsx_pdf_output_missing"),
            XLSX_PDF_TIMEOUT_MESSAGE: self._text("runtime_xlsx_pdf_timeout"),
        }
        if message in known_messages:
            return known_messages[message]

        if message.startswith(XLSX_PDF_FAILURE_PREFIX):
            detail = message.removeprefix(XLSX_PDF_FAILURE_PREFIX).strip()
            localized_message = self._text("runtime_xlsx_pdf_failed")
            if detail:
                return f"{localized_message}\n{detail}"
            return localized_message

        if message.startswith(COPILOT_FAILURE_PREFIX):
            detail = message.removeprefix(COPILOT_FAILURE_PREFIX).strip()
            localized_message = self._text("runtime_copilot_failed")
            if detail:
                return f"{localized_message}\n{detail}"
            return localized_message

        unsupported_prefix = "Unsupported file type: "
        if message.startswith(unsupported_prefix):
            return self._text(
                "validation_unsupported_type",
                suffix=message.removeprefix(unsupported_prefix),
            )

        return message

    def _text(self, key: str, **kwargs: object) -> str:
        return get_text(self._language, key, **kwargs)