from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from ._main_window import MainWindow
from ._temp_cleanup import cleanup_markitdown_temp_dirs


def main() -> int:
    cleanup_markitdown_temp_dirs()
    app = QApplication(sys.argv)
    app.setApplicationName("MarkItDown GUI")
    app.setOrganizationName("MarkItDown")
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()
    return app.exec()
