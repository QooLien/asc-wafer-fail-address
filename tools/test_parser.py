"""Compatibility entry point for the current streaming parser tests."""
import runpy
from pathlib import Path
runpy.run_path(str(Path(__file__).with_name("test_streaming_parser.py")), run_name="__main__")
