"""Module framework for the web platform.

Each automation module declares a schema (the fields the user fills in on the
web form) and an `execute` entry point that runs *inside the runner process on
the user's local machine*. The server only stores the schema and dispatches jobs.
"""

from __future__ import annotations

import abc
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

# Field types supported by the dynamic job form.
#   text|password|number|file|folder|select|checkbox|textarea|date
FIELD_TYPES = {
    "text", "password", "number", "file", "folder", "select", "checkbox",
    "textarea", "date", "hidden",
}


@dataclass
class Field:
    key: str
    label: str
    type: str = "text"
    required: bool = False
    default: Any = ""
    help: str = ""
    accept: str = ""          # e.g. ".pdf,.xlsx"
    options: List[Dict[str, Any]] = field(default_factory=list)
    placeholder: str = ""
    multiple: bool = False    # allow multiple files when type == file

    def to_dict(self) -> Dict[str, Any]:
        return {
            "key": self.key,
            "label": self.label,
            "type": self.type,
            "required": self.required,
            "default": self.default,
            "help": self.help,
            "accept": self.accept,
            "options": self.options,
            "placeholder": self.placeholder,
            "multiple": self.multiple,
        }


@dataclass
class Module:
    key: str
    name: str
    description: str
    icon: str = "📦"
    category: str = "Utilities"
    requires_browser: bool = False
    requires_credentials: bool = False
    fields: List[Field] = field(default_factory=list)
    ported: bool = False     # True once a web `execute` implementation exists
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "key": self.key,
            "name": self.name,
            "description": self.description,
            "icon": self.icon,
            "category": self.category,
            "requires_browser": self.requires_browser,
            "requires_credentials": self.requires_credentials,
            "fields": [f.to_dict() for f in self.fields],
            "ported": self.ported,
            "notes": self.notes,
        }


class ModuleContext:
    """Runtime context handed to a module's `execute` inside the runner.

    `work_dir` is a scratch directory on the *local machine*; `params` holds the
    user-submitted form values and `uploaded` maps the field key to the list of
    files the server received and the runner downloaded.
    """

    def __init__(
        self,
        job_id: int,
        title: str,
        params: Dict[str, Any],
        work_dir: Path,
        uploaded: Dict[str, List[Path]],
        credentials: Dict[str, str],
        log: Callable[[str], None],
        progress: Callable[[int, str], None],
        proxies: Optional[List[Dict[str, Any]]] = None,
    ) -> None:
        self.job_id = job_id
        self.title = title
        self.params = params
        self.work_dir = work_dir
        self.uploaded = uploaded
        self.credentials = credentials
        self._log = log
        self._progress = progress
        # Admin-managed proxies scoped to this job's runner host, e.g.
        # [{"scheme":"http","host":"...","port":80,...}]. Empty when no proxy
        # is configured for the runner's scope.
        self.proxies: List[Dict[str, Any]] = list(proxies or [])

    def proxy_url(self, index: int = 0) -> Optional[str]:
        """Return a ``scheme://host:port`` URL (no creds) for Selenium, or None."""
        if index >= len(self.proxies):
            return None
        p = self.proxies[index]
        return f"{p.get('scheme','http')}://{p.get('host')}:{p.get('port')}"

    def log(self, message: str) -> None:
        self._log(str(message))

    def progress(self, value: int, label: str = "") -> None:
        self._progress(max(0, min(100, int(value))), label)


class ModuleError(Exception):
    """Raised by a module to fail a job with a user-visible message."""


class WebModule(abc.ABC):
    """Base class for ported automation modules."""

    module: Module

    @abc.abstractmethod
    def execute(self, ctx: ModuleContext) -> None:
        """Run the automation on the local machine.

        Must call ``ctx.log`` for output and ``ctx.progress`` for updates, and
        raise :class:`ModuleError` on failure.
        """


def dumps_params(params: Dict[str, Any]) -> str:
    return json.dumps(params, default=str)


def loads_params(raw: Optional[str]) -> Dict[str, Any]:
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {}
