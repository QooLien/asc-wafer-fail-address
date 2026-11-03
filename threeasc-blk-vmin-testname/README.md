# ThreeASC Monitor / Function (v0.14.1)

Download and open templates/ASC_ThreeASC_Monitor_Function.xlsm in Windows desktop Excel. All VBA is embedded; no Python or separate VBA import is required.

## Config

- A2:C901: Monitor item (TNO / Func. / Test Voltage). Used for Vmin Judge. Default TNO350-360 W-Low pBL and TNO370-380 R-Low pBL, each with 1.20/1.035/0.99/0.90/0.81/0.765/0.72/0.68/0.64/0.60/0.56 V.
- D2:F901: Function item (TNO / Func. / Test Voltage). Initially blank. Imports Fail Address and FBC only. No Cat Fail or Vmin judgment.
- Each TNO must be unique across both lists and between 100 and 999. Blank rows are allowed. Func. and numeric voltage are required for every selected TNO.
- Q2: Vmin threshold. Blank/0 uses Legacy; a positive integer enables cumulative FBC.
- Q6:R...: optional ordered DUT X/Y list.
- J2:O4: WT source mapping; I2:I6: status; T2:Y8: BLK diagnostics.
- Config button at J13 imports exactly three ASC files. Full LOT and WNO must match; numeric STEP sorts WT1, WT2, WT3.

## Output

Original Fail Address output is retained. Its selected TNO order is Monitor list first, then Function list, with separate WT1/2/3 column groups. No Split functionality.

Each Func. has one BLK_Distribution worksheet. Columns A:F are TNO/Vmin label, Func., Test Voltage, X, Y, WT. Columns G:JB are BLK0..255, preserving raw indices. For each Chip/TNO, WT1/2/3 occupy adjacent colored rows. Missing records are Not Found, never zero.

Only Monitor TNOs contribute to Vmin. Their order follows Config. Function rows never contribute to cumulative FBC or previous voltage, even when the Func. name matches. A group with Monitor items ends each Chip with three independent WT Vmin rows. A Function-only group has no Vmin rows. No Cat Fail classification is performed.

## Validation

Tests cover original reader/parser and address/TNO output for 1/3/9/900 shuffled selections, three WT isolation, all256 raw BLK indices, first duplicates, missing and malformed data, Legacy/cumulative Vmin, dual-list input validation and address indices, Function exclusion from Vmin, Function-only output without Vmin, exact embedded VBA and Office2016 OpenXML validation. Native Excel2016 macro execution has not been tested in this environment.

Synthetic examples/default_tnos use all22 default Monitor TNOs. Older examples/three_asc fixtures require their four configured TNOs; they are used by regression tests.

## v0.14.1 recovery fix

FBC parser errors retain group/sheet/row state until BLK_Release, preventing a secondary Error424 from masking the original FBC parse diagnostic. Original error/line details remain in Config T2. A parser failure invalidates BLK data for that WT; it is never silently treated as valid zero FBC. Original Fail Address parser and Vmin rules are unchanged. Added regression coverage for the actual shipped failure handler.
