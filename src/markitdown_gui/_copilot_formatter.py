from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from ._temp_cleanup import COPILOT_TEMP_PREFIX, build_temp_dir_prefix

COPILOT_MISSING_MESSAGE = (
    "GitHub Copilot CLI executable was not found. Install GitHub Copilot CLI or set its path in the GUI."
)
COPILOT_AUTH_MESSAGE = (
    "GitHub Copilot CLI is not authenticated. Open the CLI once and complete /login, or configure a valid token."
)
COPILOT_EMPTY_OUTPUT_MESSAGE = "GitHub Copilot CLI returned no formatted Markdown."
COPILOT_TIMEOUT_MESSAGE = "GitHub Copilot CLI timed out before returning formatted Markdown."
COPILOT_FAILURE_PREFIX = "GitHub Copilot CLI failed:"
COPILOT_VERSION_CHECK_TIMEOUT_MESSAGE = (
    "GitHub Copilot CLI did not respond to the version check."
)
COPILOT_VERSION_CHECK_FAILED_MESSAGE = (
    "GitHub Copilot CLI could not be verified."
)
COPILOT_LAUNCH_FAILED_MESSAGE = "GitHub Copilot CLI could not be launched."

DEFAULT_COPILOT_PROMPT = (
    "Rewrite the referenced Markdown for readability and consistency. "
    "Return only Markdown with no code fences or commentary. "
    "Preserve the original language, facts, links, tables, and headings. "
    "Do not mention GitHub Copilot or that AI was used."
)

_ANSI_ESCAPE_RE = re.compile(r"\x1B\[[0-?]*[ -/]*[@-~]")


class CopilotCliError(RuntimeError):
    pass


@dataclass(frozen=True)
class CopilotCliProbeResult:
    resolved_command: str | None
    status: str
    detail: str = ""


def _build_hidden_process_kwargs() -> dict[str, object]:
    if os.name != "nt":
        return {}
    return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}


def detect_copilot_cli_command(*, allow_vscode_wrapper: bool = False) -> str | None:
    configured_path = os.environ.get("COPILOT_CLI_PATH", "").strip()
    if configured_path:
        return configured_path

    for candidate in _find_copilot_cli_candidates():
        if allow_vscode_wrapper or not _is_vscode_wrapper_path(candidate):
            return candidate

    if allow_vscode_wrapper:
        wrapper_path = _get_vscode_wrapper_path()
        if wrapper_path:
            return wrapper_path

    return None


def resolve_copilot_cli_command(command: str | None) -> str | None:
    configured_command = command.strip() if command else ""
    if configured_command:
        return configured_command
    return detect_copilot_cli_command()


def probe_copilot_cli_command(
    command: str | None,
    *,
    timeout_seconds: int = 10,
) -> CopilotCliProbeResult:
    resolved_command = resolve_copilot_cli_command(command)
    if not resolved_command:
        return CopilotCliProbeResult(None, "missing")

    last_detail = ""
    for arguments in (
        [resolved_command, "--binary-version"],
        [resolved_command, "version", "--silent"],
    ):
        try:
            completed = subprocess.run(
                arguments,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout_seconds,
                env=_build_command_environment(),
                check=False,
                **_build_hidden_process_kwargs(),
            )
        except FileNotFoundError:
            return CopilotCliProbeResult(resolved_command, "missing")
        except subprocess.TimeoutExpired:
            last_detail = COPILOT_VERSION_CHECK_TIMEOUT_MESSAGE
            continue

        detail = _clean_text(f"{completed.stdout}\n{completed.stderr}")
        if completed.returncode == 0:
            version_line = detail.splitlines()[0] if detail else ""
            return CopilotCliProbeResult(resolved_command, "ready", version_line)

        last_detail = detail or COPILOT_VERSION_CHECK_FAILED_MESSAGE
        if "cannot find github copilot cli" in last_detail.lower():
            return CopilotCliProbeResult(resolved_command, "missing")

    if last_detail == COPILOT_VERSION_CHECK_TIMEOUT_MESSAGE:
        return CopilotCliProbeResult(
            resolved_command,
            "timeout",
            COPILOT_VERSION_CHECK_TIMEOUT_MESSAGE,
        )

    return CopilotCliProbeResult(
        resolved_command,
        "error",
        last_detail or COPILOT_VERSION_CHECK_FAILED_MESSAGE,
    )


def launch_copilot_cli(command: str | None) -> None:
    resolved_command = resolve_copilot_cli_command(command)
    if not resolved_command:
        raise CopilotCliError(COPILOT_MISSING_MESSAGE)

    if _is_vscode_wrapper_path(resolved_command):
        wrapper_probe = probe_copilot_cli_command(resolved_command, timeout_seconds=5)
        if wrapper_probe.status != "ready":
            raise CopilotCliError(COPILOT_MISSING_MESSAGE)

    powershell_command = shutil.which("powershell.exe") or shutil.which("powershell")
    if powershell_command:
        try:
            subprocess.Popen(
                [
                    powershell_command,
                    "-NoExit",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    f'& "{resolved_command}"',
                ],
                creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
                env=_build_command_environment(),
            )
            return
        except OSError as exc:
            raise CopilotCliError(COPILOT_LAUNCH_FAILED_MESSAGE) from exc

    try:
        subprocess.Popen([resolved_command], env=_build_command_environment())
    except OSError as exc:
        raise CopilotCliError(COPILOT_LAUNCH_FAILED_MESSAGE) from exc


def _build_command_environment() -> dict[str, str]:
    environment = os.environ.copy()
    environment.setdefault("NO_COLOR", "1")
    return environment


def _find_copilot_cli_candidates() -> list[str]:
    candidates: list[str] = []

    where_command = shutil.which("where.exe") or shutil.which("where")
    if where_command:
        try:
            completed = subprocess.run(
                [where_command, "copilot"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=5,
                env=_build_command_environment(),
                check=False,
                **_build_hidden_process_kwargs(),
            )
        except (OSError, subprocess.TimeoutExpired):
            completed = None
        if completed is not None and completed.returncode == 0:
            for line in completed.stdout.splitlines():
                candidate = line.strip()
                if candidate and candidate not in candidates:
                    candidates.append(candidate)

    discovered_path = shutil.which("copilot")
    if discovered_path and discovered_path not in candidates:
        candidates.append(discovered_path)

    winget_link_path = _get_winget_link_path()
    if winget_link_path and winget_link_path not in candidates:
        candidates.append(winget_link_path)

    for package_path in _get_winget_package_paths():
        if package_path not in candidates:
            candidates.append(package_path)

    wrapper_path = _get_vscode_wrapper_path()
    if wrapper_path and wrapper_path not in candidates:
        candidates.append(wrapper_path)

    return candidates


def _get_vscode_wrapper_path() -> str | None:
    appdata = os.environ.get("APPDATA", "").strip()
    if not appdata:
        return None

    wrapper_path = (
        Path(appdata)
        / "Code"
        / "User"
        / "globalStorage"
        / "github.copilot-chat"
        / "copilotCli"
        / "copilot.bat"
    )
    if wrapper_path.exists():
        return str(wrapper_path)
    return None


def _get_winget_link_path() -> str | None:
    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if not local_app_data:
        return None

    link_path = Path(local_app_data) / "Microsoft" / "WinGet" / "Links" / "copilot.exe"
    if link_path.exists():
        return str(link_path)
    return None


def _get_winget_package_paths() -> list[str]:
    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if not local_app_data:
        return []

    package_root = Path(local_app_data) / "Microsoft" / "WinGet" / "Packages"
    if not package_root.exists():
        return []

    matches: list[Path] = []
    for package_dir in package_root.glob("GitHub.Copilot_*"):
        candidate = package_dir / "copilot.exe"
        if candidate.exists():
            matches.append(candidate)

    matches.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return [str(path) for path in matches]


def _is_vscode_wrapper_path(command: str) -> bool:
    normalized_command = Path(command).expanduser().as_posix().lower()
    return "/code/user/globalstorage/github.copilot-chat/copilotcli/copilot." in normalized_command


def _clean_text(text: str) -> str:
    return _ANSI_ESCAPE_RE.sub("", text).replace("\r\n", "\n").strip()


class CopilotCliFormatter:
    def __init__(
        self,
        command: str | None = None,
        *,
        timeout_seconds: int | None = None,
    ) -> None:
        self._command = command.strip() if command else None
        self._timeout_seconds = timeout_seconds

    def format_markdown(
        self,
        markdown: str,
        *,
        source_path: Path,
        output_path: Path,
    ) -> str:
        command = self._command or detect_copilot_cli_command()
        if not command:
            raise CopilotCliError(COPILOT_MISSING_MESSAGE)

        if _is_vscode_wrapper_path(command):
            wrapper_probe = probe_copilot_cli_command(command, timeout_seconds=5)
            if wrapper_probe.status != "ready":
                raise CopilotCliError(COPILOT_MISSING_MESSAGE)

        with tempfile.TemporaryDirectory(
            prefix=build_temp_dir_prefix(COPILOT_TEMP_PREFIX)
        ) as temp_dir:
            temp_markdown_path = Path(temp_dir) / output_path.name
            temp_markdown_path.write_text(markdown, encoding="utf-8")
            prompt = self._build_prompt(temp_markdown_path, source_path, output_path)

            try:
                completed = subprocess.run(
                    [
                        command,
                        "-p",
                        prompt,
                        "--allow-all-paths",
                        "--allow-all-tools",
                        "--no-ask-user",
                        "--silent",
                    ],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=self._timeout_seconds,
                    cwd=str(source_path.parent),
                    env=self._build_environment(),
                    check=False,
                    **_build_hidden_process_kwargs(),
                )
            except FileNotFoundError as exc:
                raise CopilotCliError(COPILOT_MISSING_MESSAGE) from exc
            except subprocess.TimeoutExpired as exc:
                raise CopilotCliError(COPILOT_TIMEOUT_MESSAGE) from exc

        if completed.returncode != 0:
            raise CopilotCliError(self._build_failure_message(completed))

        formatted_markdown = self._normalize_output(completed.stdout)
        if not formatted_markdown:
            raise CopilotCliError(COPILOT_EMPTY_OUTPUT_MESSAGE)

        return formatted_markdown

    def _build_prompt(
        self,
        markdown_path: Path,
        source_path: Path,
        output_path: Path,
    ) -> str:
        return (
            f"{DEFAULT_COPILOT_PROMPT} "
            f"Source file name: {source_path.name}. "
            f"Target markdown file name: {output_path.name}. "
            f"Input markdown: @{markdown_path}"
        )

    def _build_environment(self) -> dict[str, str]:
        return _build_command_environment()

    def _build_failure_message(self, completed: subprocess.CompletedProcess[str]) -> str:
        detail = self._clean_text(f"{completed.stderr}\n{completed.stdout}")
        detail_lower = detail.lower()
        if (
            "login" in detail_lower
            or "not authenticated" in detail_lower
            or "authentication" in detail_lower
            or "token" in detail_lower and ("invalid" in detail_lower or "expired" in detail_lower)
            or "subscription" in detail_lower and "copilot" in detail_lower
        ):
            return COPILOT_AUTH_MESSAGE
        if not detail:
            return COPILOT_FAILURE_PREFIX
        return f"{COPILOT_FAILURE_PREFIX}\n{detail}"

    def _normalize_output(self, output: str) -> str:
        cleaned_output = _clean_text(output)
        if not cleaned_output:
            return ""

        lines = cleaned_output.splitlines()
        if len(lines) >= 2 and lines[0].startswith("```") and lines[-1].strip() == "```":
            cleaned_output = "\n".join(lines[1:-1]).strip()

        return cleaned_output

    def _clean_text(self, text: str) -> str:
        return _clean_text(text)