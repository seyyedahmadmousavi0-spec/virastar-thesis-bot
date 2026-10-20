#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
if [[ ! -x .venv/bin/python ]]; then
    printf '%s\n' 'Create the project virtual environment first; see README.md.' >&2
    exit 1
fi
exec .venv/bin/python -B -X utf8 bot.py
