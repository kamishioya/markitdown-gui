__all__ = ["main"]


def main() -> int:
	from ._app import main as run_app

	return run_app()
