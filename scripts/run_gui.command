#!/bin/sh
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
PYTHONPATH="$SCRIPT_DIR/../src" python3 -m word_journal_manuscript_converter.gui
