"""Embedded GSTR-1 JSON generator entry point."""

import os
import runpy


_UI_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "json_generator_ui.py")


def render(db_path=None):
    """Render the file-based GSTR-1 JSON generator inside the main app.

    The generator's Sales Register CSV/Excel upload remains available in the
    UI, including its template, validation, preview, and JSON download flow.
    """
    runpy.run_path(_UI_FILE, run_name="__main__")
