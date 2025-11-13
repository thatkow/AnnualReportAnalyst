"""Shared logging utilities for the Annual Report Analyst application."""
from __future__ import annotations

import logging

LOGGER_NAME = "annualreport"


def get_logger() -> logging.Logger:
    """Return the shared application logger instance."""
    return logging.getLogger(LOGGER_NAME)
