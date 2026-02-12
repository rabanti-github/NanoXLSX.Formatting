![NanoXLSX](https://raw.githubusercontent.com/rabanti-github/NanoXLSX.Formatting/refs/heads/main/Documentation/NanoXLSXlib.png) 

# NanoXLSX.Formatting

![NuGet Version](https://img.shields.io/nuget/v/NanoXLSX.Formatting)
![NuGet Downloads](https://img.shields.io/nuget/dt/NanoXLSX.Formatting)
![GitHub License](https://img.shields.io/github/license/rabanti-github/NanoXLSX.Formatting)

NanoXLSX is a small .NET library written in C#, to create and read Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

---

The **Formatting** package is responsible to enable NanoXLSX for the handling of **structured text** within a text cell (inline formatting).

Currently supported:

- Adding text runs (paragraphs) by a Builder or manually
- Defining Font styles per run
- Adding phonetic information (important for East Asian languages)
- Reading and writing runs, phonetic information and inline Font styles 

---

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX.Formatting/blob/master/Changelog.md)** for recent updates.

## What's new in version 3.x

This is the first release if this package. It was set to v 3.x, to be consistent with the NanoXLSX v3 ecosystem

## Road map

Planned features:

- Style builder for general cell styles
- Assistant for valid custom format codes
- Support for conditional cell formatting

## Requirements

[NanoXLSX.Formatting](https://www.nuget.org/packages/NanoXLSX.Formatting) is not intended as standalone package. It requires **[NanoXLSX.Core](https://www.nuget.org/packages/NanoXLSX.Core)**, and is normally part of the meta-package **[NanoXLSX](https://www.nuget.org/packages/NanoXLSX)**

**You find all technical requirements in the main repository: [NanoXLSX](https://github.com/rabanti-github/NanoXLSX)**

### General requirements

- .NET 4.5 or newer / .NET Standard
- NanoXLSX.Core as only dependency

### Utility dependencies

The Test project and GitHub Actions may also require dependencies like unit testing frameworks or workflow steps. However, **none of these dependencies are essential to build the library**. They are just utilities. The test dependencies ensure efficient unit testing and code coverage. The GitHub Actions dependencies are used for the automatization of releases and API documentation

## Installation

### Using NuGet

By package Manager (PM):

```sh
Install-Package NanoXLSX.Formatting
```

By .NET CLI:

```sh
dotnet add package NanoXLSX.Formatting
```

## Usage

### Quick Start (manual)

```c#
Workbook workbook = new Workbook("sheet1");                          // Create new workbook
FormattedText formattedText = new FormattedText();                   // Create new formatted text
Font strikeFont = new Font() { Strike = true };                      // Create a new font style
formattedText.AddRun("strike", strikeFont);                          // Add text and style as run to the formatted text
formattedText.AddLineBreak();                                        // Add a line break after the first run
Font boldFont = new Font() { Bold = true, Name = "Tahoma" };         // Create a second font
formattedText.AddRun("bold", boldFont);                              // Add a second run to the formatted text
workbook.CurrentWorksheet.AddFormattedTextCell(formattedText, "A1"); // Add formatted text to the worksheet
```

### Quick Start (builders)

```c#
Workbook workbook = new Workbook("sheet1");                            // Create new workbook
InlineStyleBuilder inlineStyleBuilder = new InlineStyleBuilder()       // Create a new inline style builder
    .Color("FFAABBCC")                                                 // Add properties to the builder
    .Italic()
    .Size(18);
FormattedTextBuilder builder = new FormattedTextBuilder();             // Create a new formatted text builder
builder.AddRun("朝日", inlineStyleBuilder.Build());                     // Add a text and style to the builder
builder.AddPhoneticRun("あさひ", 0, 2);                                 // Add a phonetic run  to the builder
builder.AddRun("Asahi", new Font() { VerticalAlign = Font.VerticalTextAlignValue.Superscript }); // Add another text and style to the builder
workbook.CurrentWorksheet.AddFormattedTextCell(builder.Build(), 0, 0); // Create the formatted text and add it to the cell                                     // Save the workbook as myWorkbook.xlsx
```

## Further References

- See the full **package API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX.Formatting/](https://rabanti-github.github.io/NanoXLSX.Formatting/).
- See the full NanoXLSX **API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX/](https://rabanti-github.github.io/NanoXLSX/).

## License

NanoXLSX.Formatting is published under the **MIT** license.

The project / package is developed with as much compliance to this license as only possible.
Please visit the main repository [NanoXLSX](https://github.com/rabanti-github/NanoXLSX) for compliance a scan, provided by Fossa 
