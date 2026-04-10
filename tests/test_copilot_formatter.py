from __future__ import annotations

import os
import unittest
from unittest import mock
from pathlib import Path

from markitdown_gui._copilot_formatter import (
    COPILOT_MISSING_MESSAGE,
    CopilotCliError,
    CopilotCliFormatter,
    CopilotCliProbeResult,
    detect_copilot_cli_command,
    launch_copilot_cli,
)


class CopilotFormatterDetectionTests(unittest.TestCase):
    def test_detect_prefers_winget_package_path(self) -> None:
        winget_path = (
            r"C:\Users\example\AppData\Local\Microsoft\WinGet\Packages"
            r"\GitHub.Copilot_Microsoft.Winget.Source_8wekyb3d8bbwe\copilot.exe"
        )

        with mock.patch(
            "markitdown_gui._copilot_formatter._find_copilot_cli_candidates",
            return_value=[winget_path],
        ):
            detected = detect_copilot_cli_command()

        self.assertEqual(detected, winget_path)

    def test_detect_skips_vscode_wrapper_when_real_cli_exists(self) -> None:
        wrapper_path = (
            r"C:\Users\example\AppData\Roaming\Code\User\globalStorage"
            r"\github.copilot-chat\copilotCli\copilot.bat"
        )
        real_path = r"C:\Program Files\GitHub Copilot\copilot.exe"

        with mock.patch(
            "markitdown_gui._copilot_formatter._find_copilot_cli_candidates",
            return_value=[wrapper_path, real_path],
        ):
            detected = detect_copilot_cli_command()

        self.assertEqual(detected, real_path)

    def test_detect_returns_none_when_only_vscode_wrapper_exists(self) -> None:
        wrapper_path = (
            r"C:\Users\example\AppData\Roaming\Code\User\globalStorage"
            r"\github.copilot-chat\copilotCli\copilot.bat"
        )

        with mock.patch(
            "markitdown_gui._copilot_formatter._find_copilot_cli_candidates",
            return_value=[wrapper_path],
        ):
            detected = detect_copilot_cli_command()

        self.assertIsNone(detected)


class CopilotFormatterLaunchTests(unittest.TestCase):
    def test_launch_rejects_unverified_vscode_wrapper(self) -> None:
        wrapper_path = (
            r"C:\Users\example\AppData\Roaming\Code\User\globalStorage"
            r"\github.copilot-chat\copilotCli\copilot.bat"
        )

        with mock.patch(
            "markitdown_gui._copilot_formatter.probe_copilot_cli_command",
            return_value=CopilotCliProbeResult(wrapper_path, "missing"),
        ):
            with mock.patch("subprocess.Popen") as popen_mock:
                with self.assertRaisesRegex(CopilotCliError, COPILOT_MISSING_MESSAGE):
                    launch_copilot_cli(wrapper_path)

        popen_mock.assert_not_called()

    def test_formatter_rejects_unverified_vscode_wrapper(self) -> None:
        wrapper_path = (
            r"C:\Users\example\AppData\Roaming\Code\User\globalStorage"
            r"\github.copilot-chat\copilotCli\copilot.bat"
        )
        formatter = CopilotCliFormatter(command=wrapper_path)

        with mock.patch(
            "markitdown_gui._copilot_formatter.probe_copilot_cli_command",
            return_value=CopilotCliProbeResult(wrapper_path, "missing"),
        ):
            with self.assertRaisesRegex(CopilotCliError, COPILOT_MISSING_MESSAGE):
                formatter.format_markdown(
                    "# Sample\n",
                    source_path=Path(r"C:\work\sample.xlsx"),
                    output_path=Path(r"C:\work\sample.md"),
                )

    def test_formatter_uses_noninteractive_auto_approval_and_hidden_process(self) -> None:
        formatter = CopilotCliFormatter(command=r"C:\Tools\copilot.exe", timeout_seconds=30)
        completed = subprocess_completed_process(
            stdout="# Sample\n\nHello\n",
            stderr="",
            returncode=0,
        )

        with mock.patch(
            "markitdown_gui._copilot_formatter.subprocess.run",
            return_value=completed,
        ) as run_mock:
            result = formatter.format_markdown(
                "# Sample\n\nHello\n",
                source_path=Path(r"C:\work\sample.xlsx"),
                output_path=Path(r"C:\work\sample.md"),
            )

        self.assertEqual(result, "# Sample\n\nHello")
        command_args = run_mock.call_args.args[0]
        self.assertIn("--allow-all-tools", command_args)
        self.assertIn("--no-ask-user", command_args)
        self.assertIn("--allow-all-paths", command_args)
        self.assertEqual(run_mock.call_args.kwargs.get("timeout"), 30)
        if os.name == "nt":
            self.assertEqual(
                run_mock.call_args.kwargs.get("creationflags"),
                getattr(__import__("subprocess"), "CREATE_NO_WINDOW", 0),
            )

    def test_formatter_defaults_to_unlimited_timeout(self) -> None:
        formatter = CopilotCliFormatter(command=r"C:\Tools\copilot.exe")
        completed = subprocess_completed_process(
            stdout="# Sample\n\nHello\n",
            stderr="",
            returncode=0,
        )

        with mock.patch(
            "markitdown_gui._copilot_formatter.subprocess.run",
            return_value=completed,
        ) as run_mock:
            formatter.format_markdown(
                "# Sample\n\nHello\n",
                source_path=Path(r"C:\work\sample.xlsx"),
                output_path=Path(r"C:\work\sample.md"),
            )

        self.assertIsNone(run_mock.call_args.kwargs.get("timeout"))


def subprocess_completed_process(*, stdout: str, stderr: str, returncode: int):
    import subprocess

    return subprocess.CompletedProcess(
        args=["copilot"],
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
    )
