# v0.17.0 - Func/voltage headers and adjacent WT comparison

- Fail Address Row1 keeps WT labels and displays Func._TestVoltage instead of TNO, e.g. W-Low pBL_1.20v, W-Low pBL_1.035v, W-Low pBL_0.765v.
- Use Config B/C and E/F for names and voltages. Missing optional metadata falls back to numeric TNO. Address indices and Fail flags remain unchanged. Header height150 pt; body26 pt.
- FBC rows are grouped by Chip X/Y, configured TNO order, then WT1/2/3. Whole256-FBC rows and temperature colors are preserved. No summation or fabricated missing-WT zeros.
- Retain the independent J13 Address-only and J16 FBC-only buttons.

Validation passed: actual VBA label helpers and shuffled header-to-TNO mapping, original address1/3/9/900-TNO regression, WT sort keys and rank helper including one-row case, action isolation, embedded source checks and Office2016 file-format validation. The Excel native sort and macros have not been executed in native Excel2016, and the user's ASC has not been tested here.
