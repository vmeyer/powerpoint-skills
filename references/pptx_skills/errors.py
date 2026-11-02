"""Exceptions shared by all commands."""


class PptxSkillsError(Exception):
    """A user-facing error; the CLI prints it as `error: <message>` and exits 2."""
