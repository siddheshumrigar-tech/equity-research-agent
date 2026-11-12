"""Typed errors with the process exit codes the CLI returns to the calling agent."""


class EquityAgentError(Exception):
    """Base class. ``exit_code`` is what ``generate_equity_report.py`` exits with."""

    exit_code = 1


class SymbolNotFound(EquityAgentError):
    exit_code = 2


class BankNotSupported(EquityAgentError):
    exit_code = 3


class DataValidationError(EquityAgentError):
    exit_code = 4


class NetworkError(EquityAgentError):
    exit_code = 5


class InputValidationError(EquityAgentError):
    """analyst_inputs.json or narrative.json written by the host agent failed validation."""

    exit_code = 6
