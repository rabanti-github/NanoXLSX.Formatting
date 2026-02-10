# Change Log - NanoXLSX.Formatting

## v3.0.0

---
Release Date: **10.02.2026** <sup>(DMY)</sup>

- Initial release of the formatting library
- Support of reading and writing inline formatting of text cells
- Exposing of the data types:
  - `FormattedText` (Contains runs and styles)
  - `FormattedTextBuilder` (Utility to build FormattedText instances)
  - `InlineStyleBuilder` (Utility to build a Font style component, applied to a text run)
  - `TextRun` (Contains a text and a possible style; paragraph)
  - `PhoneticRun` (Contains phonetic information e.g. for East Asian languages)
  - `PhoneticProperties` (Contains options, how phonetic runs are displayed)

Note *I*: If a cell contains inline formatting, such a cell will be automatically hold a `FormattedText`  object instead of `string`, when a workbook is read.

Note *II*: The version of the package is set to 3.0.0  instead 1.0.0, to be consistent with the NanoXLSX v3 ecosystem  