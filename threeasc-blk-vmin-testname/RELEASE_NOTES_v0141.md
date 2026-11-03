# v0.14.1 - Preserve FBC parser error diagnostics

Fix Error424 during BLK Vmin WT1: the FBC error handler incorrectly released grouping dictionaries needed for output. Cleanup now runs only in BLK_Release. Original FBC error details and line numbers remain available in Config T2. FBC errors no longer mask the original issue with a secondary object-required error.

Actual failure-handler regression, shared-stream parser, Monitor-only Vmin, embedded-source checks and Office2016 OpenXML validation passed. Native Excel2016 macro execution and the user ASC have not been tested here.
