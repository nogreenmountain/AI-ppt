"""Compatibility shim for the report CLI.

The CLI implementation historically lives at ``slide2pptx.report_cli``
(package level). Tests, however, import it as
``slide2pptx.report.report_cli`` (i.e. inside the ``report`` sub-package).
This module re-exports the public surface expected by those tests so that
``from slide2pptx.report import report_cli`` keeps working without
duplicating the implementation.

Public attributes expected by the test suite:

* :func:`main` - the ``__main__`` entry point.
* :data:`EXIT_OK`, :data:`EXIT_INPUT`, :data:`EXIT_RENDER` - exit codes.
* :mod:`renderer` - the renderer sub-module (``SlideRenderer``,
  :class:`RendererError`, :func:`_resolve_powershell`).
* :mod:`models` - the dataclass sub-module (e.g. :class:`RenderResult`).
* :data:`LOG` - the package logger.
"""

from __future__ import annotations

# Re-export the entire top-level report_cli module.  This single ``import *``
# is sufficient because the upstream ``report_cli`` already exposes every
# name the tests touch (``main``, ``EXIT_*``, ``renderer``, ``models``,
# ``LOG``) at the module level - including the imported ``renderer`` and
# ``models`` sub-modules that tests reach for via attribute access
# (e.g. ``report_cli.renderer.SlideRenderer``).
from slide2pptx.report_cli import (  # noqa: F401,F403
    EXIT_INPUT,
    EXIT_OK,
    EXIT_OTHER,
    EXIT_RENDER,
    LOG,
    main,
    models,
    renderer,
)

__all__ = [
    "EXIT_INPUT",
    "EXIT_OK",
    "EXIT_OTHER",
    "EXIT_RENDER",
    "LOG",
    "main",
    "models",
    "renderer",
]
