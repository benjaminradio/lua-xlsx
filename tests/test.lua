-- Usage: lua tests/test.lua [cpath-pattern] [fixture]
--   e.g. lua tests/test.lua "./target/release/lib?.so"
if arg[1] then package.cpath = arg[1] .. ";" .. package.cpath end
local fixture = arg[2] or "tests/fixtures/sample.xlsx"
local xlsx = require("xlsx")

local passed = 0
local function eq(got, want, msg)
  if got ~= want then
    error(("%s: expected %s, got %s"):format(msg or "assertion", tostring(want), tostring(got)), 2)
  end
  passed = passed + 1
end
local function fails(fn, pattern, msg)
  local ok, e = pcall(fn)
  if ok then error((msg or "expected an error") .. " but call succeeded", 2) end
  if pattern and not tostring(e):find(pattern, 1, true) then
    error(("%s: error %q does not contain %q"):format(msg or "error", tostring(e), pattern), 2)
  end
  passed = passed + 1
end

-- helpers ------------------------------------------------------------------
eq(xlsx.colname(1), "A"); eq(xlsx.colname(26), "Z"); eq(xlsx.colname(27), "AA")
eq(xlsx.colname(16384), "XFD")
eq(xlsx.colnumber("ab"), 28); eq(xlsx.colnumber("XFD"), 16384)
eq(xlsx.address(7, 2), "B7"); eq(xlsx.address(7, 2, "r1c1"), "R7C2")
local r, c = xlsx.parseref("$C$9"); eq(r, 9); eq(c, 3)
r, c = xlsx.parseref("r12c30"); eq(r, 12); eq(c, 30)
fails(function() xlsx.parseref("A0") end, "invalid")
fails(function() xlsx.parseref("XFE1") end, "invalid")
fails(function() xlsx.colname(0) end)

-- workbook -----------------------------------------------------------------
local wb = xlsx.open(fixture)
eq(#wb, 4)
local names = wb:sheetnames()
eq(names[1], "People"); eq(names[4], "Hidden")
local info = wb:sheetinfo()
eq(info[1].visibility, "visible"); eq(info[4].visibility, "hidden")
eq(wb:definednames().TaxRate, "People!$B$8")
fails(function() wb:sheet("Nope") end, "no sheet named")
fails(function() wb:sheet(9) end, "out of range")
fails(function() xlsx.open("/nonexistent.xlsx") end, "cannot open")
fails(function() xlsx.load("not a zip") end, "cannot read workbook")

-- raw access: data[row][col], absolute from A1 -----------------------------
local s = wb:sheet("People")
eq(s.name, "People")
eq(s[1][1], "Quarterly Report")
eq(s[3][1], "Name"); eq(s[4][1], "Alice"); eq(s[4][2], 30)
eq(math.type and math.type(s[4][2]) or "integer", "integer")
eq(s[5][2], 25.5)
eq(s[6][2], nil)                       -- empty cell
eq(s[4][4], "2020-01-15")              -- date -> ISO string
eq(s[5][4], "2021-06-30T13:30:00")     -- datetime
eq(s[4][6], true); eq(s[5][6], false)  -- booleans
eq(s[4][8], "#DIV/0!")                 -- error cell
eq(s[99], nil); eq(s[0], nil); eq(s[4][99], nil)
eq(#s, 8)
eq(s.dimension, "A1:H8")

-- A1 and R1C1 notation (absolute) -------------------------------------------
eq(s["A4"], "Alice"); eq(s["R4C1"], "Alice"); eq(s["r4c2"], 30)
eq(s:cell("B4"), 30); eq(s:cell("R4C2"), 30); eq(s:cell(4, 2), 30)
eq(s:cell("$B$4"), 30)
eq(s:cell("Z99"), nil)
fails(function() s:cell("bogus") end, "invalid cell reference")
fails(function() s:cell(0, 1) end)
local sub = s["A3:D6"]
eq(#sub, 4); eq(sub[1][1], "Name"); eq(sub[2][1], "Alice"); eq(sub.region, "A3:D6")
eq(s:range("R3C1:R6C4")[2][3], "Paris")
eq(s:range(3, 1, 6, 4)[4][3], "Berlin")

-- headers from the first row of a region ------------------------------------
local p = wb:sheet("People", { headers = true, range = "A3:D6" })
eq(#p, 3)
eq(p.headers[1], "Name"); eq(p.headers[4], "Joined")
eq(p[1].Name, "Alice"); eq(p[1].Age, 30); eq(p[1][1], "Alice")
eq(p[2].City, "London"); eq(p[3].Age, nil); eq(p[3].City, "Berlin")
eq(p[1].Nope, nil)
eq(#p[1], 4)
-- A1 notation is still absolute on a headed view
eq(p["A4"], "Alice"); eq(p:cell("C6"), "Berlin")
-- open-ended range
local p2 = wb:sheet("People", { headers = true, range = "A3:" })
eq(#p2, 5); eq(p2[1].Name, "Alice")
-- manual headers
local m = wb:sheet("People", { headers = { "who", "years", "town", "since" }, range = "A4:D6" })
eq(#m, 3); eq(m[1].who, "Alice"); eq(m[2].town, "London"); eq(m[3].since, "2022-12-01")
-- with_headers is non-destructive
local plain = s:range("A3:D6")
local h = plain:with_headers(true)
eq(#h, 3); eq(h[1].Name, "Alice")
eq(#h:with_headers(false), 4)
eq(#plain:with_headers({ "a", "b", "c", "d" }), 4)
eq(plain:with_headers({ "a", "b", "c", "d" })[1].a, "Name")

-- iteration ----------------------------------------------------------------
local names_seen = {}
for i, row in ipairs(p) do names_seen[i] = row.Name end
eq(table.concat(names_seen, ","), "Alice,Bob,Carol")
names_seen = {}
for i, row in pairs(p) do names_seen[#names_seen + 1] = i .. row.Name end
eq(table.concat(names_seen, ","), "1Alice,2Bob,3Carol")
local n = 0
for row in p:rows() do n = n + 1 end
eq(n, 3)
local keys = {}
for k, v in pairs(p[1]) do keys[#keys + 1] = k end
table.sort(keys)
eq(table.concat(keys, ","), "Age,City,Joined,Name")
keys = {}
for k, v in pairs(p[3]) do keys[#keys + 1] = k end   -- Age is empty -> skipped
table.sort(keys)
eq(table.concat(keys, ","), "City,Joined,Name")

-- plain tables for Lus queries / tocsv --------------------------------------
local t = p:totable()
eq(type(t), "table"); eq(#t, 3); eq(t[1].Name, "Alice"); eq(t[2].Age, 25.5); eq(t[3].Age, nil)
eq(t[1][1], nil)                       -- header-keyed like fromcsv(.., true)
local raw = plain:totable()
eq(raw[1][1], "Name"); eq(raw[2][2], 30)
eq(p:col("Age")[1], 30); eq(p:col(3)[2], "London")
fails(function() p:col("zzz") end, "no column")
eq(p:row(2).Name, "Bob"); eq(p:row(9), nil)

-- merged cells --------------------------------------------------------------
local merged = wb:sheet("People"):merged()
eq(merged[1], "A1:D1")
local plainm = wb:sheet("Messy")
eq(plainm["A2"], "North"); eq(plainm["A3"], nil)
local filled = wb:sheet("Messy", { fill_merged = true, headers = true, range = "A1:C6" })
eq(filled[1].Region, "North"); eq(filled[2].Region, "North"); eq(filled[3].Region, "North")
eq(filled[4].Region, "South"); eq(filled[5].Region, "South"); eq(filled[5].Item, "nail")
local wbf = xlsx.open(fixture, { fill_merged = true })
eq(wbf:sheet("Messy")["A4"], "North")

-- duplicate / empty headers -------------------------------------------------
local d = wb:sheet("Messy", { headers = true, range = "B10:E11" })
eq(d.headers[1], "Name"); eq(d.headers[2], "Name_2"); eq(d.headers[3], ""); eq(d.headers[4], "Name_3")
eq(d[1].Name, 1); eq(d[1].Name_2, 2); eq(d[1].Name_3, 4); eq(d[1][3], 3)
local dt = d:totable()[1]
eq(dt.Name, 1); eq(dt.Name_2, 2); eq(dt[3], 3); eq(dt.Name_3, 4)

-- find ----------------------------------------------------------------------
local fr, fc = s:find("Total")
eq(fr, 8); eq(fc, 1)
eq(s:cell(fr, fc + 1) == nil, true)   -- formula has no cached value
eq(s:find("nothing here"), nil)
local fr2 = s:find("lic", { contains = true })
eq(fr2, 4)
eq(s:find("ALICE", { ignorecase = true }), 4)
eq(s:find("^Car", { pattern = true }), 6)
eq((s:find(25.5)), 5)
local all = s:findall("Name")
eq(#all, 1); eq(all[1].ref, "A3"); eq(all[1].value, "Name")
eq(#wb:sheet("People"):findall(true), 1)

-- formulas ------------------------------------------------------------------
eq(s:formula("G4"), "B4+1")
eq(s:formula("B8"), "SUM(B4:B6)")
eq(s:formula("A4"), nil)

-- dates / errors options ----------------------------------------------------
local serial = wb:sheet("People", { dates = "serial" })
eq(serial[4][4], 43845.0)
eq(wb:sheet("People", { errors = "nil" })[4][8], nil)
fails(function() wb:sheet("People", { dates = "bogus" }) end, "bad option")
local wbs = xlsx.open(fixture, { dates = "serial" })
eq(wbs:sheet("People")[4][4], 43845.0)

-- tables --------------------------------------------------------------------
local tables = wb:tables()
eq(#tables, 1); eq(tables[1], "Sales")
local sales = wb:table("Sales")
eq(#sales, 3); eq(sales.headers[2], "Units")
eq(sales[1].Region, "EU"); eq(sales[3].Price, 3.0); eq(sales[2].Units, 7)
eq(sales["A2"], "EU")                 -- absolute coordinates still work
fails(function() wb:table("Nope") end)

-- convenience indexing and tostring -----------------------------------------
eq(wb[1].name, "People"); eq(wb["Tbl"].name, "Tbl"); eq(wb["nothing"], nil)
eq(tostring(p):find("xlsx.sheet", 1, true), 1)

-- load from memory ----------------------------------------------------------
local f = assert(io.open(fixture, "rb"))
local bytes = f:read("*a"); f:close()
eq(xlsx.load(bytes):sheet("People")[4][1], "Alice")

print(("ok: %d assertions passed"):format(passed))
