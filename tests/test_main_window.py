from __future__ import annotations

import os
import tempfile
import unittest
from unittest import mock
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from markitdown_gui._copilot_setup_dialog import CopilotSetupDialog
from markitdown_gui._main_window import MainWindow


class MainWindowLocalizationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QApplication.instance() or QApplication([])

    def _create_settings(self) -> tuple[tempfile.TemporaryDirectory[str], QSettings]:
        temp_dir = tempfile.TemporaryDirectory(prefix="markitdown-gui-settings-")
        settings = QSettings(
            os.path.join(temp_dir.name, "settings.ini"),
            QSettings.Format.IniFormat,
        )
        settings.clear()
        return temp_dir, settings

    def test_default_language_is_japanese(self) -> None:
        temp_dir, settings = self._create_settings()
        self.addCleanup(temp_dir.cleanup)

        window = MainWindow(settings=settings, default_copilot_command="")
        self.addCleanup(window.close)

        self.assertEqual(window._language, "ja")
        self.assertEqual(window._language_combo.currentData(), "ja")
        self.assertEqual(window._language_label.text(), "表示言語")
        self.assertEqual(window._convert_button.text(), "変換開始")
        self.assertIn("MarkItDown 出力", window._output_dir_edit.text())
        self.assertEqual(
            window._copilot_checkbox.text(),
            "GitHub Copilot CLI で Markdown を整形する",
        )
        self.assertEqual(window._settings_menu.title(), "設定")
        self.assertEqual(
            window._copilot_setup_action.text(),
            "GitHub Copilot CLI セットアップ",
        )

    def test_language_selection_is_persisted(self) -> None:
        temp_dir, settings = self._create_settings()
        self.addCleanup(temp_dir.cleanup)

        first_window = MainWindow(settings=settings, default_copilot_command="")
        self.addCleanup(first_window.close)

        english_index = first_window._language_combo.findData("en")
        self.assertGreaterEqual(english_index, 0)
        first_window._language_combo.setCurrentIndex(english_index)

        self.assertEqual(first_window._language, "en")
        self.assertEqual(first_window._language_label.text(), "Display Language")
        self.assertEqual(first_window._convert_button.text(), "Convert")
        self.assertIn("MarkItDown Output", first_window._output_dir_edit.text())

        first_window._copilot_checkbox.setChecked(True)
        first_window._copilot_command_edit.setText(r"C:\Tools\copilot.exe")

        settings.sync()

        second_window = MainWindow(settings=settings, default_copilot_command="")
        self.addCleanup(second_window.close)

        self.assertEqual(second_window._language, "en")
        self.assertEqual(second_window._language_combo.currentData(), "en")
        self.assertEqual(second_window._language_label.text(), "Display Language")
        self.assertTrue(second_window._copilot_checkbox.isChecked())
        self.assertEqual(second_window._copilot_command_edit.text(), r"C:\Tools\copilot.exe")
        self.assertEqual(second_window._settings_menu.title(), "Settings")
        self.assertEqual(
            second_window._copilot_setup_action.text(),
            "GitHub Copilot CLI Setup",
        )

    def test_copilot_setup_dialog_updates_main_window(self) -> None:
        temp_dir, settings = self._create_settings()
        self.addCleanup(temp_dir.cleanup)

        window = MainWindow(settings=settings, default_copilot_command="")
        self.addCleanup(window.close)

        class StubDialog:
            DialogCode = type("DialogCode", (), {"Accepted": 1, "Rejected": 0})

            def __init__(self, **kwargs) -> None:
                self.kwargs = kwargs

            def exec(self) -> int:
                return self.DialogCode.Accepted

            def copilot_enabled(self) -> bool:
                return True

            def copilot_command(self) -> str:
                return r"C:\Guided\copilot.exe"

        with mock.patch("markitdown_gui._main_window.CopilotSetupDialog", StubDialog):
            window._open_copilot_setup_dialog()

        self.assertTrue(window._copilot_checkbox.isChecked())
        self.assertEqual(window._copilot_command_edit.text(), r"C:\Guided\copilot.exe")
        self.assertEqual(
            settings.value("copilot/command", "", type=str),
            r"C:\Guided\copilot.exe",
        )

    def test_stage_progress_updates_status_bar_and_progress_bar(self) -> None:
        temp_dir, settings = self._create_settings()
        self.addCleanup(temp_dir.cleanup)

        source_path = Path(temp_dir.name) / "sample.pdf"
        source_path.write_bytes(b"%PDF-1.4")

        window = MainWindow(settings=settings, default_copilot_command="")
        self.addCleanup(window.close)
        window._add_paths([str(source_path)])
        window._set_busy_state(True)

        window._on_stage_changed(str(source_path), 1, 2, 42, "markitdown")

        self.assertEqual(window._status_progress_bar.value(), 42)
        self.assertIn("42%", window.statusBar().currentMessage())
        self.assertIn("Markdown", window.statusBar().currentMessage())


class CopilotSetupDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QApplication.instance() or QApplication([])

    def test_auto_detect_uses_default_command(self) -> None:
        dialog = CopilotSetupDialog(
            locale="ja",
            copilot_enabled=False,
            copilot_command="",
            default_copilot_command=r"C:\Tools\copilot.exe",
        )
        self.addCleanup(dialog.close)

        dialog._apply_detected_command()

        self.assertEqual(dialog.copilot_command(), r"C:\Tools\copilot.exe")
        self.assertIn(r"C:\Tools\copilot.exe", dialog._status_value_label.text())


if __name__ == "__main__":
    unittest.main()