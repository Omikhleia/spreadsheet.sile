# spreadsheet.sile

[![License](https://img.shields.io/github/license/Omikhleia/spreadsheet.sile?label=License)](LICENSE)
[![Luacheck](https://img.shields.io/github/actions/workflow/status/Omikhleia/spreadsheet.sile/luacheck.yml?branch=main&label=Luacheck&logo=Lua)](https://github.com/Omikhleia/spreadsheet.sile/actions?workflow=Luacheck)
[![Luarocks](https://img.shields.io/luarocks/v/Omikhleia/spreadsheet.sile?label=Luarocks&logo=Lua)](https://luarocks.org/modules/Omikhleia/spreadsheet.sile)

This collection of packages for the [SILE](https://github.com/sile-typesetter/sile) typesetting system provides (experimental) support for reading and processing spreadsheets, and for rendering them as tables in SILE-generated documents.

The supported formats are:

- OpenDocument Spreadsheet Document (`.ods`).
- Office Open XML Workbook (`.xlsx`).

## Installation

These packages require SILE v0.14 or upper.

Installation relies on the **luarocks** package manager.

To install the latest development version and all its dependencies (see below),
you may use the provided “rockspec”:

```
luarocks --lua-version 5.4 install --dev spreadsheet.sile
```

(Adapt to your version of Lua, if need be, and refer to the SILE manual for more
detailed 3rd-party package installation information.)

The collection depends on [ptable.sile](https://github.com/Omikhleia/ptable.sile) which is also installed when using the above procedure.

## Usage

Once the collection is installed, the **spreadsheet** experimental package is available.
It provides basic tools for parsing and processing spreadsheet documents, converting them to tables in output documents.

E.g. in SIL language:

```
\use[module=packages.spreadsheet]
\spreadsheet[src=<filename>, sheet=<name|number>, header=<boolean>]
```

The command reads an external spreadsheet file and generates a table.

The `sheet` option defaults to 1, but may be set to another index or an explicit sheet name, for spreadsheet documents containing several sheets.

The `header` option (false by default) specifies whether first row should be considered as a header row or not in the generated table.

Please note that the spreadsheets are expected to be reasonably simple: merged cells or rows are not supported, content styling is ignored, etc.

## Status

This is sort of a Proof-of-Concept showing how the Lua and SILE ecosystems can be hacked in interesting, though probably unintended, ways.

The current implementation is not very efficient — it doesn't cache anything, for instance.
It's not very flexible either — for instance, all columns have the same width and the table is made full-line wide.

You might have better results converting your spreadsheets via other means, to either SIL language or to anything else suitable (e.g. Markdown "pipe tables" for inclusion via [markdown.sile](https://github.com/Omikhleia/markdown.sile), etc.) — Though for quickly including a table in a document, perhaps you'll find this package worth a try?

In other terms, its author himself is not that sure this solution is the right way to go.
Who knows, though, what you might do with it?

## License

The code in this repository is released under the MIT License, Copyright 2023, Omikhleia.

Some ZIP-processing code is derived and adapted from LuaRocks, also released under the MIT License, Copyright 2007-2011, Kepler Project, Copyright 2011-2022, the LuaRocks project authors.
