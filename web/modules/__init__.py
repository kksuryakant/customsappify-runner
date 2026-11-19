"""Runner-side module framework (schemas + implementations).

Present so ``web.modules`` is a regular package when the feature bundle is
extracted on a runner: Nuitka freezes ``web`` into the executable, and the
per-job bundle only patches ``web.__path__``. Without this file ``web.modules``
resolves to a namespace package that does not expose the extracted ``impl``
subpackage, breaking ``web.modules.impl`` imports on the runner.
"""
