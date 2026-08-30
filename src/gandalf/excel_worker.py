"""Entry point for the Windows VM Microsoft Excel worker service."""

from gandalf.excel_mcp import worker_main


def main() -> None:
    """Run the worker service."""
    worker_main()


if __name__ == "__main__":
    main()
