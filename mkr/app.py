"""Application entry."""
from __future__ import annotations

import sys

from mkr.gui import launch


def main() -> int:
    # Force Python 3.11+ to shut up if someone runs on older versions.
    if sys.version_info < (3, 11):
        sys.stderr.write("ClaraEasyExcel requires Python 3.11 or later.\n")
        return 1
    return launch()


if __name__ == "__main__":
    raise SystemExit(main())
