# v0.18.0 - Stop Sign endpoint Fail Maps

- Config G1 Stop Sign is red/white; G Pattern Name selects the same-row Function TNO in D/E/F as the cumulative endpoint. Monitor A/B/C never participates in Map Judge.
- Count only endpoint BLKs with FBC>0, never re-add earlier TNOs. Green0 Pass, red1..256 Fail BLK Count, gray No Data, yellow Data Error.
- Raw paired X/Y inventory includes DUT/DC/ISB/no-FBC records and is shared across WT1/2/3. Map geometry is data-driven, with one three-WT sheet per Pattern.
- Add Chip Locations and BLK Fail Count outputs; J19 rebuilds Maps without Raw reimport. J13 remains Address-only; J16 imports FBC and builds Maps.
- Fix native VBA EOF call: explicit Call FinishRecord prevents a line-start label from dropping the last record. Native tests cover both complete and incomplete final records.
- Passed native Excel16 tests for synthetic64/128 Chips, two Patterns, three WTs, endpoint501-only0/3/256 counts, missing/DC-only gray, parser-error yellow, managed Map cleanup and Address preservation. Also passed original Address regression and Office2016 format checks. Company ASC/environment remain untested.
