from __future__ import annotations

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from ._copilot_formatter import (
    CopilotCliError,
    COPILOT_LAUNCH_FAILED_MESSAGE,
    detect_copilot_cli_command,
    launch_copilot_cli,
    probe_copilot_cli_command,
)
from ._localization import get_text

INSTALL_GUIDE_URL = "https://docs.github.com/en/copilot/how-tos/set-up/install-copilot-cli"


class CopilotSetupDialog(QDialog):
    def __init__(
        self,
        *,
        locale: str,
        copilot_enabled: bool,
        copilot_command: str,
        default_copilot_command: str,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._language = locale
        self._default_copilot_command = default_copilot_command

        self.setModal(True)
        self.resize(760, 560)

        self._build_ui()
        self._copilot_checkbox.setChecked(copilot_enabled)
        self._command_edit.setText(copilot_command or default_copilot_command)
        self._status_value_label.setText(self._text("copilot_setup_status_not_checked"))
        self._apply_language()

    def copilot_enabled(self) -> bool:
        return self._copilot_checkbox.isChecked()

    def copilot_command(self) -> str:
        return self._command_edit.text().strip()

    def _build_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        self._intro_label = QLabel(self)
        self._intro_label.setWordWrap(True)
        main_layout.addWidget(self._intro_label)

        self._settings_group = QGroupBox(self)
        settings_layout = QVBoxLayout(self._settings_group)

        self._copilot_checkbox = QCheckBox(self)
        settings_layout.addWidget(self._copilot_checkbox)

        form_layout = QFormLayout()
        self._command_label = QLabel(self)
        self._command_edit = QLineEdit(self)
        form_layout.addRow(self._command_label, self._command_edit)
        settings_layout.addLayout(form_layout)

        action_row = QHBoxLayout()
        self._auto_detect_button = QPushButton(self)
        self._browse_button = QPushButton(self)
        self._check_button = QPushButton(self)
        self._launch_button = QPushButton(self)
        self._docs_button = QPushButton(self)
        action_row.addWidget(self._auto_detect_button)
        action_row.addWidget(self._browse_button)
        action_row.addWidget(self._check_button)
        action_row.addWidget(self._launch_button)
        action_row.addWidget(self._docs_button)
        settings_layout.addLayout(action_row)

        self._status_label = QLabel(self)
        self._status_value_label = QLabel(self)
        self._status_value_label.setWordWrap(True)
        status_layout = QFormLayout()
        status_layout.addRow(self._status_label, self._status_value_label)
        settings_layout.addLayout(status_layout)

        main_layout.addWidget(self._settings_group)

        self._steps_group = QGroupBox(self)
        steps_layout = QVBoxLayout(self._steps_group)
        self._steps_label = QLabel(self)
        self._steps_label.setWordWrap(True)
        self._steps_edit = QPlainTextEdit(self)
        self._steps_edit.setReadOnly(True)
        steps_layout.addWidget(self._steps_label)
        steps_layout.addWidget(self._steps_edit)
        main_layout.addWidget(self._steps_group, stretch=1)

        self._button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        self._button_box.accepted.connect(self.accept)
        self._button_box.rejected.connect(self.reject)
        main_layout.addWidget(self._button_box)

        self._auto_detect_button.clicked.connect(self._apply_detected_command)
        self._browse_button.clicked.connect(self._pick_copilot_command)
        self._check_button.clicked.connect(self._check_command)
        self._launch_button.clicked.connect(self._launch_cli)
        self._docs_button.clicked.connect(self._open_install_guide)

    def _apply_language(self) -> None:
        self.setWindowTitle(self._text("copilot_setup_dialog_title"))
        self._intro_label.setText(self._text("copilot_setup_intro"))
        self._settings_group.setTitle(self._text("copilot_setup_settings_group"))
        self._copilot_checkbox.setText(self._text("checkbox_copilot_format"))
        self._command_label.setText(self._text("copilot_command_label"))
        self._command_edit.setPlaceholderText(self._text("copilot_command_placeholder"))
        self._auto_detect_button.setText(self._text("copilot_setup_button_auto_detect"))
        self._browse_button.setText(self._text("button_browse"))
        self._check_button.setText(self._text("copilot_setup_button_check"))
        self._launch_button.setText(self._text("copilot_setup_button_launch"))
        self._docs_button.setText(self._text("copilot_setup_button_open_docs"))
        self._status_label.setText(self._text("copilot_setup_status_label"))
        self._steps_group.setTitle(self._text("copilot_setup_steps_group"))
        self._steps_label.setText(self._text("copilot_setup_steps_label"))
        self._steps_edit.setPlainText(self._text("copilot_setup_steps_text"))

    def _pick_copilot_command(self) -> None:
        current = self._command_edit.text().strip() or self._default_copilot_command or ""
        command_path, _ = QFileDialog.getOpenFileName(
            self,
            self._text("dialog_select_copilot_command"),
            current,
            self._text("dialog_copilot_command_filter"),
        )
        if command_path:
            self._command_edit.setText(command_path)
            self._status_value_label.setText(self._text("copilot_setup_status_not_checked"))

    def _apply_detected_command(self) -> None:
        detected_command = self._default_copilot_command or detect_copilot_cli_command()
        self._command_edit.setText(detected_command or "")
        if detected_command:
            self._status_value_label.setText(
                self._text("copilot_setup_status_detected", path=detected_command)
            )
            return
        self._status_value_label.setText(self._text("copilot_setup_status_missing"))

    def _check_command(self) -> None:
        result = probe_copilot_cli_command(self._command_edit.text().strip())
        if result.status == "ready":
            if result.detail:
                self._status_value_label.setText(
                    self._text("copilot_setup_status_ready", detail=result.detail)
                )
            else:
                self._status_value_label.setText(
                    self._text("copilot_setup_status_ready_no_detail")
                )
            return
        if result.status == "missing":
            self._status_value_label.setText(self._text("copilot_setup_status_missing"))
            return
        if result.status == "timeout":
            self._status_value_label.setText(self._text("copilot_setup_status_timeout"))
            return

        detail = result.detail.strip()
        if detail:
            self._status_value_label.setText(
                self._text("copilot_setup_status_error", detail=detail)
            )
            return
        self._status_value_label.setText(self._text("copilot_setup_status_error_no_detail"))

    def _launch_cli(self) -> None:
        try:
            launch_copilot_cli(self._command_edit.text().strip())
        except CopilotCliError as exc:
            QMessageBox.warning(
                self,
                self._text("copilot_setup_launch_title"),
                self._localize_runtime_message(str(exc)),
            )
            return

        QMessageBox.information(
            self,
            self._text("copilot_setup_launch_title"),
            self._text("copilot_setup_launch_message"),
        )

    def _open_install_guide(self) -> None:
        QDesktopServices.openUrl(QUrl(INSTALL_GUIDE_URL))

    def _localize_runtime_message(self, message: str) -> str:
        known_messages = {
            "GitHub Copilot CLI executable was not found. Install GitHub Copilot CLI or set its path in the GUI.": self._text("runtime_copilot_missing"),
            COPILOT_LAUNCH_FAILED_MESSAGE: self._text("copilot_setup_launch_failed"),
        }
        return known_messages.get(message, message)

    def _text(self, key: str, **kwargs: object) -> str:
        return get_text(self._language, key, **kwargs)