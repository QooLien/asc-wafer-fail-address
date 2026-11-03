# v0.15.0 - Single-sheet WT1/2/3 FBC

- Keep the original validated Fail Address output.
- Consolidate WT1/WT2/WT3 into one FBC by BLK worksheet: WT, TNO, X, Y, BLK0..255.
- Remove Func. grouping and Vmin/Cat Fail judgment. Only Config A/D TNO columns are required.
- Skip TNO0/unselected FBC records, correcting the reported TNO must be positive failure.
- Preserve exact256-value validation and first-record-wins duplicate handling for selected TNOs.
- Remove obsolete workbooks and tools from main; retain wafer coordinates and necessary address regression baselines.

Tests passed: original address mapping for1/3/9/900 selected TNOs, all256 raw FBC positions, TNO0 skip, three-file isolation, malformed selected FBC diagnostics, and the actual batch output loop over3078 rows. Embedded-source checks and Office2016 OpenXML validation passed. Native Excel2016 execution and user ASC validation are not available.
