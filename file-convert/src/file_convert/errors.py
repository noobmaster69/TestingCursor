"""Typed errors mapped to CLI exit codes."""

from __future__ import annotations


class ConvertError(Exception):
    """Base error with exit code."""

    exit_code: int = 1

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


class UserError(ConvertError):
    exit_code = 1


class DependencyError(ConvertError):
    exit_code = 2


class ConversionError(ConvertError):
    exit_code = 3


class ValidationError(UserError):
    pass
