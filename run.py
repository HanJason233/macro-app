import os
import sys

# Some Windows/Python environments set DPI awareness before Qt initializes.
# Use a compatible Windows DPI level so Qt won't attempt PMv2 and log Access Denied.
if sys.platform == "win32":
    os.environ.setdefault("QT_QPA_PLATFORM", "windows:dpiawareness=1")

from macro_app.app import main


if __name__ == "__main__":
    raise SystemExit(main())
