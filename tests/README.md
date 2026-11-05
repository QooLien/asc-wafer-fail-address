# Native validation

Validation JSON records summarize completed desktop Excel/PowerPoint tests. BLK voltage tests use the synthetic64Chip input from the release package, never production ASC.

`expected_blk_voltage.py` derives2304 expected BLK states/first voltages/WR counts from the original synthetic ASC and Config CSV. PowerShell tests require desktop Excel and an editable project-root path. `test_blk_voltage_events.ps1` uses a test-only event wrapper; it is not embedded in the release workbooks. These scripts retain the development workspace paths for reproducibility and may require adapting the root/build fixture location.

Excel users do not need Python or these tests to run the shipped XLSM.
