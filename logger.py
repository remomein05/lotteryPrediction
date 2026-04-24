"""Logger class that tees stdout to a markdown file."""

import sys
from typing import Any, cast


class Logger:
    """Logs terminal output to a markdown file while writing to stdout."""

    def __init__(self, filename: str) -> None:
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding="utf-8")
        self.log.write("# Terminal Execution Log\n```text\n")

    def write(self, message: str) -> None:
        """Write message to both terminal and log file."""
        self.terminal.write(message)
        if not self.log.closed:
            self.log.write(message)

    def flush(self) -> None:
        """Flush both terminal and log file."""
        self.terminal.flush()
        if not self.log.closed:
            self.log.flush()

    def close_log(self) -> None:
        """Close the log file."""
        if not self.log.closed:
            self.log.write("\n```\n")
            self.log.close()


def setup_logging(log_filename: str = "result.md") -> None:
    """Setup stdout/stderr logging to a markdown file."""
    import builtins

    # Replace stdout with logger
    sys.stdout = Logger(log_filename)  # type: ignore

    # Hook input to capture user keystrokes
    original_input = builtins.input

    def logged_input(prompt: object = "") -> str:
        user_val = str(original_input(prompt))
        out = cast(Any, sys.stdout)
        if hasattr(out, "log") and out.log and not out.log.closed:
            out.log.write(user_val + "\n")
            out.log.flush()
        return user_val

    builtins.input = cast(Any, logged_input)
