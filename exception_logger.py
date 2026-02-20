"""Configures logging for uncaught exceptions."""

import logging
import sys


class ExtraNewlineFormatter(logging.Formatter):
    """Wraps logging.Formatter to add extra newline after log entries."""

    def format(self, record: logging.LogRecord) -> str:
        """Adds extra newline after log entries."""
        original = super().format(record)
        return f"{original}\n"


def configure_logger() -> None:
    """Sets up the root logger for logging uncaught exceptions."""
    logger = logging.getLogger()
    logger.setLevel(logging.ERROR)  # disable warnings from pypdf

    # set up file handler
    file_handler = logging.FileHandler("exceptions.log", mode="a")
    formatter = ExtraNewlineFormatter("%(asctime)s")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    def handle_exception(exc_type, exc_value, exc_traceback):
        """Wraps uncaught exceptions so they can be logged."""
        logger.error("", exc_info=(exc_type, exc_value, exc_traceback))
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

    # overwrite internal function for uncaught exceptions
    sys.excepthook = handle_exception
