"""Entry point of the `pptx` command; subcommands live in COMMAND_MODULES."""

import argparse
import importlib
import json
import sys
from types import ModuleType

from .errors import PptxSkillsError

OK, FINDINGS, ERROR = 0, 1, 2

COMMAND_MODULES: list[ModuleType] = []

# Command modules, imported on first use so that importing cli stays cheap and cycle-free.
_COMMAND_NAMES = ("inventory",)


def _load_commands() -> None:
    for name in _COMMAND_NAMES:
        module = importlib.import_module(f"{__package__}.{name}")
        if module not in COMMAND_MODULES:
            COMMAND_MODULES.append(module)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pptx", description="Design and check PowerPoint decks.")
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    _load_commands()
    for module in COMMAND_MODULES:
        module.register(subparsers)
    return parser


def run_handler(handler, args) -> int:
    try:
        return handler(args)
    except PptxSkillsError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return ERROR


def emit(args, data: dict, text: str) -> None:
    if getattr(args, "json", False):
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        print(text)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return exc.code if isinstance(exc.code, int) else ERROR
    handler = getattr(args, "handler", None)
    if handler is None:
        parser.print_help(sys.stderr)
        return ERROR
    return run_handler(handler, args)
