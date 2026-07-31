"""Project-level pytest configuration.

Adjusts Python path so ``import slide2pptx`` resolves from the
``python/`` sub-tree, regardless of where pytest is launched.
"""

import sys
from pathlib import Path


def _ensure_python_path() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    python_root = repo_root / "python"
    sys.path.insert(0, str(python_root))


_ensure_python_path()
