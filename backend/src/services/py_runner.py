#!/usr/bin/env python3
import sys
import json
import importlib
import traceback
import os

# Directory with user scripts (default from user's workspace). Override with PY_SCRIPTS_DIR env var.
DEFAULT_SCRIPTS_DIR = r"C:\Users\estatistica007\Documents\nextProjects\RelatoriosContabeis\scripts"
# Directory where Python scripts should write their output. Can be overridden with PY_OUTPUT_DIR env var.
DEFAULT_OUTPUT_DIR = r"C:\Users\estatistica007\Documents\nextProjects\RelatoriosContabeis\backend\output"

scripts_dir = os.environ.get('PY_SCRIPTS_DIR', DEFAULT_SCRIPTS_DIR)
output_dir = os.environ.get('PY_OUTPUT_DIR', DEFAULT_OUTPUT_DIR)

# Ensure scripts imports still work even if we chdir to output_dir
if scripts_dir and os.path.isdir(scripts_dir):
    sys.path.insert(0, scripts_dir)

# Ensure output directory exists and set it as current working dir so
# any files created with relative paths are placed into the output folder.
try:
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    os.chdir(output_dir)
except Exception:
    # best-effort; fall back to scripts_dir if chdir fails
    try:
        if scripts_dir and os.path.isdir(scripts_dir):
            os.chdir(scripts_dir)
    except Exception:
        pass

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: py_runner.py <module> <function>"}))
        sys.exit(2)
    module_name = sys.argv[1]
    func_name = sys.argv[2]
    try:
        data = json.load(sys.stdin) if not sys.stdin.isatty() else {}
    except Exception:
        data = {}

    try:
        mod = importlib.import_module(module_name)
        func = getattr(mod, func_name)
        # call with kwargs if dict
        if isinstance(data, dict):
            result = func(**data)
        elif isinstance(data, list):
            result = func(*data)
        else:
            result = func(data)
        # try to jsonify result
        try:
            print(json.dumps(result, default=str))
        except Exception:
            # if not serializable, return a simple success message
            print(json.dumps({"message": str(result)}))
    except Exception as e:
        traceback.print_exc()
        try:
            print(json.dumps({"error": str(e)}))
        except Exception:
            print(json.dumps({"error": "Unknown error"}))
        sys.exit(1)

if __name__ == '__main__':
    main()
