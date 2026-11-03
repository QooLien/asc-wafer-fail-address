# v0.16.0 - Independent Address / FBC buttons

- Config J13: Address-only import. Config J16: FBC-only import.
- FBC-only mode neither parses addresses nor allocates address flags nor writes the Fail Address worksheet. Address-only mode does not parse or write FBC.
- Rebuild FBC parsing around each actual FBC BY BLK record, taking TNO/X/Y from the same header line. General metadata cannot trigger the old global Malformed header check or supply stale coordinates.
- Support normal, tab, full-width and non-breaking whitespace; preserve exact256 values and raw BLK0..255 order.
- Keep WT1/WT2/WT3 in one FBC sheet, with no Func. grouping or Vmin/Cat Fail judgment.
- All VBA is embedded in the single XLSM. Existing address parsing helpers and output loop are preserved.

Validation passed: separate action branches, both Forms button macro targets/anchors, supplied TNO350/DUT9/X10/Y17 header under five separator types and three chunk sizes, all256 raw indices, metadata/missing-context/malformed-record cases,3078 actual batch output rows and original address1/3/9/900-TNO regression. Embedded-source and Office2016 OpenXML checks passed. Native Excel2016 execution and the user's actual ASC remain untested.
