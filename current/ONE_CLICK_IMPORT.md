# Single-click import,1-3 ASC files

Config H2 / ImportASCFailAddresses now runs Address and FBC from the same read stream, then comparison, Maps and Map-scoped Address count coverage, followed by SaveAs and elapsed time. The original address tokenizer, numeric validation and address-to-TNO bit index are unchanged.

Inputs are a prefix of the test sequence:1 file=WT1;2 files sorted by numericSTEP=WT1/WT2;3 files=WT1/WT2/WT3. Full LOT/WNO consistency and unique STEP checks remain. Only imported phases appear in Address columns, comparison columns, Count rows and Map panels. Missing records within an imported phase retain the existing gray/unavailable behavior.

Config H:N excess cell borders are cleared. Stop Sign G border is retained. Existing button bounds and settings are preserved. Chip Locations is penultimate; Log is last, including after manual coverage refresh.

Native Excel tests cover1/2/3 files, all exact Address flag cells, expected FBC/Map/coverage counts, source index consistency, absent-WT omission, repeated3-to1 import removing stale panels/columns, and final worksheet order. Legacy1/3/9/900-TNO address regression and voltage labels also pass.
