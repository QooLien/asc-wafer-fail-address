# VBA build tooling

The embedded VBA project was generated with kay-ws/ovba-writer, using its p1_compiled.bin test fixture as the source of Excel project/reference metadata. Sample event handlers and sample standard/class modules were removed. The new container contains this project's ImportFailAddresses source and empty Excel document modules only.

Source: https://github.com/kay-ws/ovba-writer

Build revision: c2e085be54cd92f75efd52291d6f6b2ceea34c81

The tool is MIT licensed. The full upstream license is retained in `OVBA_WRITER_LICENSE.txt`.

The XLSX sheet content was authored with the bundled @oai/artifact-tool. `package_xlsm.py` only adds macro-related OPC metadata and the VBA container; it does not author spreadsheet cell content.
