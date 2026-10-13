# lua-xlsx

A native Lua module (Rust, [mlua](https://github.com/mlua-rs/mlua) +
[calamine](https://github.com/tafia/calamine)) for **reading** Excel `.xlsx`
workbooks. It is aimed at ETL over complex, semi-structured, messy
spreadsheets: banners above the data, merged label cells, several tables on one
sheet, data spread over many sheets. Lua pulls the pieces together into the
tabular shape that tools such as Power Query expect.

The API follows the conventions of [Lus](https://lus.dev)'s built-in
[`fromcsv`](https://github.com/lus-lang/lus/blob/main/lus-site/src/manual/acquis-24.mdx)
(rows of values, optional headers, `data[row].header`), and sheets can be handed
to Lus [queries](https://github.com/lus-lang/lus/blob/main/lus-site/src/manual/acquis-34.mdx)
via `sheet:totable()`. It is an ordinary Lua C module, so it also works with
stock Lua.

```lua
local xlsx = require("xlsx")
local wb   = xlsx.open("report.xlsx")

local data = wb:sheet("Sales", { headers = true, range = "A3:" })
print(#data, data[1].Region, data[1][2])      -- data[row_n].header / data[row_n][col_n]
print(data["B7"], data["R7C2"], data:cell(7, 2))   -- A1 / R1C1 / row,col (absolute)
```

## Install

Download the archive matching your Lua version and platform from the
[Releases](../../releases) page and put `xlsx.dll` / `xlsx.so` where
`package.cpath` finds it (for example next to your script, with
`package.cpath = "./?.so;./?.dll;" .. package.cpath`).

| Asset | For |
|-------|-----|
| `lua-xlsx-lua5.X-windows-x64.zip` | 64-bit [LuaBinaries](https://luabinaries.sourceforge.net/download.html) `lua5X.dll` |
| `lua-xlsx-lua5.X-windows-x86.zip` | 32-bit LuaBinaries |
| `lua-xlsx-lua5.X-linux-x64.tar.gz` | Linux x86-64 (built on Ubuntu 22.04, glibc 2.35+) |

`X` is 1–5. The Lua version must match the interpreter. On Windows the module
imports its Lua API from `lua51.dll` … `lua55.dll` by name, so the interpreter
must load that DLL (the LuaBinaries `lua5X.exe` does). The Linux module resolves the Lua API from the
host executable, so the interpreter must export its symbols (distro `lua5.x`
packages do; this has not been checked against the LuaBinaries Linux builds).

### Build from source

```sh
cargo build --release                                        # Lua 5.4 (default)
cargo build --release --no-default-features --features lua53 # lua51 … lua55
cp target/release/libxlsx.so ./xlsx.so                       # xlsx.dll on Windows
lua tests/test.lua "./?.so"
```

To publish a release, push a tag: `git tag v0.1.0 && git push origin v0.1.0`.
The *Release* workflow builds Linux x64 and Windows x64/x86 binaries for Lua
5.1–5.5 and attaches them, with checksums, to a GitHub release.

## Reading data

```lua
local wb = xlsx.open(path [, opts])      -- from a file
local wb = xlsx.load(bytes [, opts])     -- from a string already in memory
```

Workbook `opts` (also accepted per sheet): `dates`, `errors`, `fill_merged` — see
[Values](#values).

| Workbook | |
|---|---|
| `wb:sheetnames()` | array of names |
| `wb:sheetinfo()` | array of `{name=, visibility="visible"\|"hidden"\|"veryhidden"}` |
| `wb:sheet(name_or_index [, opts])` | a sheet (below) |
| `wb[1]`, `wb["Name"]`, `#wb` | shorthand for `wb:sheet(..)` with defaults; sheet count |
| `wb:tables()` / `wb:table(name [, opts])` | Excel *tables*; the sheet is the table body with the table's column names as headers |
| `wb:definednames()` | `{ Name = "Sheet1!$A$1:$B$2", ... }` |

### Sheets

Sheet options: `headers`, `range` (plus `dates`, `errors`, `fill_merged`).

`data[row_n][col_n]` uses **1-based** indexes relative to the sheet's *window*.
By default the window is the whole worksheet starting at `A1`, so `row_n` is the
Excel row number. A cell reference (`"B7"`, `"R7C2"`) is **always absolute**,
whatever window the sheet was opened with.

```lua
data[3][2]              -- row 3, column 2 of the window
#data                   -- number of data rows;  data.nrows / data.ncols
data["B7"]              -- A1 notation   (also "$B$7")
data["R7C2"]            -- R1C1 notation
data:cell("B7")         -- same, as a method;  data:cell("R7C2");  data:cell(7, 2)
data["B3:F20"]          -- a sub-sheet (window) over a range
data:range("B3:F20", { headers = true })
data:range(3, 2, 20, 6)
```

Empty cells are `nil`; use `#row` for a row's width and never rely on `ipairs`
to stop at the last cell of a row.

### Headers

`headers` is `true` (the first row of the window holds the names and is removed
from the data), or a table of names (set manually; no row is removed), or `false`.

```lua
local d = wb:sheet("People", { headers = true, range = "A3:D6" })
d[1].Name      -- by header
d[1][1]        -- numeric indexes still work
d.headers      -- { "Name", "Age", "City", "Joined" }

local m = wb:sheet("People", { headers = { "who", "years", "town", "since" }, range = "A4:D6" })
local d2 = raw:with_headers(true)            -- re-head an existing window; never mutates `raw`
```

Header names read from a row are trimmed. Duplicates become `Name`, `Name_2`, `Name_3`; empty
headers have no name and are reachable by number only. `range` accepts `"A3:D6"`,
`"R3C1:R6C4"` and the open-ended `"A3:"` (to the end of the used range).

### Iterating

```lua
for i, row in ipairs(data) do ... end    -- Lua 5.3+ (ipairs honours __index)
for i, row in pairs(data)  do ... end    -- Lua 5.2+
for row in data:rows()     do ... end    -- every version
for key, v in pairs(row)   do ... end    -- non-empty cells; key = header, else column number
```

### Plain tables, columns, queries

```lua
data:totable()      -- array of plain rows. With headers, rows are keyed by header
                    -- only (like fromcsv(s, true)); without, by column number.
data:col("Age")     -- array of one column's values; data:col(2) by position
data:row(3)         -- same as data[3]
```

`totable()` gives plain arrays of records, the shape Lus's in-memory query provider
accepts (`@rows as r where r.Age > 18 select r.Name`). For `tocsv`, which wants
sequential rows, use a window without headers: `tocsv(data:with_headers(false):totable())`.

### Messy-data helpers

```lua
data:find("Total")                       -- row, col  (absolute) of the first match, or nil
data:find("tot", { contains = true, ignorecase = true })
data:find("^Q%d", { pattern = true })    -- Lua pattern
data:findall(25.5)                       -- { {row=, col=, ref="B5", value=}, ... }
data:merged()                            -- { "A1:D1", "A2:A4", ... }
data:formula("B8")                       -- "SUM(B4:B6)" (no leading "=") or nil
data.dimension                           -- used range of the worksheet, e.g. "A1:H8"
data.region                              -- this window, e.g. "A3:D6"

local total_row, col = data:find("Total")
local total = data:cell(total_row, col + 1)   -- read relative to a label
```

`fill_merged = true` copies the top-left value of each merged region into the
other cells, so vertically merged group labels (`Region` spanning several rows)
become usable per-row values.

Module helpers: `xlsx.colname(28)` → `"AB"`, `xlsx.colnumber("AB")` → `28`,
`xlsx.address(7, 2 [, "r1c1"])` → `"B7"`/`"R7C2"`, `xlsx.parseref("B7")` → `7, 2`.

### Values

| Excel cell | Lua value |
|---|---|
| text | string |
| number | integer when whole (`30`), otherwise float (`25.5`) |
| boolean | boolean |
| empty | `nil` |
| date / datetime | ISO string: `"2024-01-31"`, `"2024-01-31T13:30:00"`, time-only `"13:30:00"` (durations `"h:mm:ss"`); `dates = "serial"` gives the Excel serial number |
| error (`#N/A`, `#DIV/0!` …) | the error text; `errors = "nil"` gives `nil` |

Formulas return their cached result (as written by Excel); a file produced by a
library that does not compute formulas has no cached value, so the cell is `nil`.

## Notes and limitations

* Read-only; `.xlsx` only (via calamine's `Xlsx` reader).
* Each `wb:sheet()` call parses that worksheet; keep and reuse the returned sheet.
* Integer and float distinction relies on Lua 5.3+; on 5.1/5.2 all numbers are doubles.
* The test suite (`tests/test.lua`, run with a `lua` interpreter) passes on Lua 5.4 and
  5.5 builds in development; CI runs it on 5.3 and 5.4. Release builds for other
  versions and for Windows are compiled but not executed by CI. Nothing has been run
  against a Lus binary yet. The module only uses the standard Lua C API, so it should
  load there, but that is untested; the Lus query and `tocsv` interop described above
  follows the Lus documentation rather than a test.
