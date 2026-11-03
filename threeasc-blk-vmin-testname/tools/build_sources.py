from pathlib import Path
import runpy
runpy.run_path(str(Path(__file__).with_name("upgrade_testnames.py")), run_name="__main__")
runpy.run_path(str(Path(__file__).with_name("upgrade_item_groups.py")), run_name="__main__")
