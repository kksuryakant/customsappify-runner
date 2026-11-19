"""BOE Tracker module (ported, no functionality).

The desktop ``BOE_Tracker`` source was an empty placeholder -- it only rendered
a "source file was empty and could not be recovered" message and contained no
tracking logic. This port mirrors that: it runs cleanly and reports that the
feature is unavailable pending source recovery.
"""

from __future__ import annotations

import logging

from ..base import Module, ModuleContext, WebModule

logger = logging.getLogger(__name__)

MODULE = Module(
    key="BOE_Tracker",
    name="BOE Tracker",
    description="Bill of Entry tracking (source module was empty in the desktop build).",
    icon="⚠️",
    category="Utilities",
    requires_browser=False,
    ported=True,
    notes="Desktop source was empty; no tracking logic available.",
    fields=[],
)


class BOETrackerModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        ctx.progress(50, "Checking availability")
        ctx.log(
            "BOE Tracker source was empty in the desktop build and could not be "
            "recovered. No tracking logic is available."
        )
        ctx.log("Please restore BOE_Tracker.py from a backup to enable this feature.")
        ctx.progress(100, "Completed")
