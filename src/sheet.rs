//! Sheet and row userdata: the `data[row][col]` / `data[row].header` API.

use std::cell::{OnceCell, RefCell};
use std::collections::HashMap;
use std::io::Cursor;
use std::rc::Rc;

use calamine::{Data, Dimensions, Range, Reader, Xlsx};
use mlua::{
    Function, Lua, MultiValue, Result, Table, UserData, UserDataFields, UserDataMethods, Value,
};

use crate::convert::{display, to_value, Fmt};
use crate::refs::{self, Pos};

fn err<S: std::fmt::Display>(msg: S) -> mlua::Error {
    mlua::Error::runtime(msg)
}

/// An open workbook, shared by every sheet loaded from it.
pub struct Wb {
    pub xlsx: RefCell<Xlsx<Cursor<Vec<u8>>>>,
    pub fmt: Fmt,
    pub fill_merged: bool,
    pub tables_loaded: std::cell::Cell<bool>,
}

/// Parsed cell data for one worksheet. Coordinates are absolute and 1-based.
pub struct Grid {
    pub name: String,
    pub range: Range<Data>,
    pub fmt: Fmt,
    pub wb: Rc<Wb>,
    formulas: OnceCell<Range<String>>,
    merged: OnceCell<Vec<Dimensions>>,
}

impl Grid {
    pub fn load(wb: &Rc<Wb>, name: &str, fmt: Fmt, fill_merged: bool) -> Result<Rc<Grid>> {
        let mut range = wb.xlsx.borrow_mut().worksheet_range(name).map_err(err)?;
        let grid_merged = OnceCell::new();
        if fill_merged {
            let merges = wb
                .xlsx
                .borrow_mut()
                .merge_cells_by_sheet_name(name)
                .map_err(err)?;
            if let (Some(rs), Some(re)) = (range.start(), range.end()) {
                for m in &merges {
                    let value = match range.get_value(m.start) {
                        Some(v) if *v != Data::Empty => v.clone(),
                        _ => continue,
                    };
                    for r in m.start.0.max(rs.0)..=m.end.0.min(re.0) {
                        for c in m.start.1.max(rs.1)..=m.end.1.min(re.1) {
                            if (r, c) != m.start {
                                range.set_value((r, c), value.clone());
                            }
                        }
                    }
                }
            }
            let _ = grid_merged.set(merges);
        }
        Ok(Rc::new(Grid {
            name: name.to_string(),
            range,
            fmt,
            wb: wb.clone(),
            formulas: OnceCell::new(),
            merged: grid_merged,
        }))
    }

    pub fn cell(&self, row: u32, col: u32) -> Option<&Data> {
        if row == 0 || col == 0 {
            return None;
        }
        self.range.get_value((row - 1, col - 1))
    }

    /// Bottom-right corner of the used range (1-based), or (0, 0) when empty.
    pub fn used_end(&self) -> (u32, u32) {
        self.range.end().map_or((0, 0), |e| (e.0 + 1, e.1 + 1))
    }

    fn formula(&self, row: u32, col: u32) -> Result<Option<String>> {
        if self.formulas.get().is_none() {
            let r = self
                .wb
                .xlsx
                .borrow_mut()
                .worksheet_formula(&self.name)
                .map_err(err)?;
            let _ = self.formulas.set(r);
        }
        let r = self.formulas.get().expect("initialised above");
        Ok(row
            .checked_sub(1)
            .zip(col.checked_sub(1))
            .and_then(|p| r.get_value(p))
            .filter(|f| !f.is_empty())
            .cloned())
    }

    fn merged(&self) -> Result<&Vec<Dimensions>> {
        if self.merged.get().is_none() {
            let m = self
                .wb
                .xlsx
                .borrow_mut()
                .merge_cells_by_sheet_name(&self.name)
                .map_err(err)?;
            let _ = self.merged.set(m);
        }
        Ok(self.merged.get().expect("initialised above"))
    }
}

/// Column names of a view, de-duplicated so every key is unique.
pub struct Headers {
    pub names: Vec<String>,
    index: HashMap<String, u32>,
}

impl Headers {
    fn new(raw: Vec<String>) -> Headers {
        let mut index = HashMap::new();
        let mut names = Vec::with_capacity(raw.len());
        for (i, name) in raw.into_iter().enumerate() {
            let mut unique = name.clone();
            if !name.is_empty() {
                let mut n = 1;
                while index.contains_key(&unique) {
                    n += 1;
                    unique = format!("{name}_{n}");
                }
                index.insert(unique.clone(), i as u32 + 1);
            }
            names.push(unique);
        }
        Headers { names, index }
    }
}

pub enum HeaderSpec {
    None,
    FromRow,
    Names(Vec<String>),
}

impl HeaderSpec {
    pub fn from_value(v: Value) -> Result<HeaderSpec> {
        match v {
            Value::Nil | Value::Boolean(false) => Ok(HeaderSpec::None),
            Value::Boolean(true) => Ok(HeaderSpec::FromRow),
            Value::Table(t) => {
                let mut names = Vec::new();
                for item in t.sequence_values::<Value>() {
                    names.push(match item? {
                        Value::String(s) => s.to_string_lossy().to_string(),
                        Value::Integer(i) => i.to_string(),
                        Value::Number(n) => n.to_string(),
                        Value::Boolean(b) => b.to_string(),
                        Value::Nil => String::new(),
                        other => {
                            return Err(err(format!(
                                "headers table may only hold strings or numbers, got {}",
                                other.type_name()
                            )))
                        }
                    });
                }
                Ok(HeaderSpec::Names(names))
            }
            other => Err(err(format!(
                "headers must be true, false or a table of names, got {}",
                other.type_name()
            ))),
        }
    }
}

/// A rectangular window onto a [`Grid`], optionally with column headers.
///
/// `data[i][j]` is relative to the window (after the header row, if the
/// headers were taken from it); `cell()` / `"A1"` access is always absolute.
pub struct SheetInner {
    pub grid: Rc<Grid>,
    /// Full region including any header row (absolute, 1-based, inclusive).
    pub r1: u32,
    pub c1: u32,
    pub r2: u32,
    pub c2: u32,
    headers: Option<Headers>,
    header_consumed: bool,
}

impl SheetInner {
    pub fn new(grid: Rc<Grid>, region: (u32, u32, u32, u32), spec: HeaderSpec) -> SheetInner {
        let (r1, c1, r2, c2) = region;
        let (headers, header_consumed) = match spec {
            HeaderSpec::None => (None, false),
            HeaderSpec::Names(n) => (Some(Headers::new(n)), false),
            HeaderSpec::FromRow => {
                let names = (c1..=c2)
                    .map(|c| {
                        grid.cell(r1, c)
                            .map(|d| display(d, grid.fmt))
                            .unwrap_or_default()
                    })
                    .map(|s| s.trim().to_string())
                    .collect();
                (Some(Headers::new(names)), r2 >= r1)
            }
        };
        SheetInner {
            grid,
            r1,
            c1,
            r2,
            c2,
            headers,
            header_consumed,
        }
    }

    fn row0(&self) -> u32 {
        self.r1 + self.header_consumed as u32
    }
    pub fn nrows(&self) -> u32 {
        (self.r2 + 1).saturating_sub(self.row0())
    }
    pub fn ncols(&self) -> u32 {
        (self.c2 + 1).saturating_sub(self.c1)
    }

    fn data(&self, i: u32, j: u32) -> Option<&Data> {
        if i < 1 || j < 1 || i > self.nrows() || j > self.ncols() {
            return None;
        }
        self.grid.cell(self.row0() + i - 1, self.c1 + j - 1)
    }

    fn value(&self, lua: &Lua, i: u32, j: u32) -> Result<Value> {
        match self.data(i, j) {
            Some(d) => to_value(lua, d, self.grid.fmt),
            None => Ok(Value::Nil),
        }
    }

    fn key_value(&self, lua: &Lua, j: u32) -> Result<Value> {
        match self
            .headers
            .as_ref()
            .and_then(|h| h.names.get(j as usize - 1))
        {
            Some(n) if !n.is_empty() => Ok(Value::String(lua.create_string(n)?)),
            _ => Ok(Value::Integer(j as mlua::Integer)),
        }
    }

    fn row_table(&self, lua: &Lua, i: u32) -> Result<Table> {
        let t = lua.create_table()?;
        for j in 1..=self.ncols() {
            let v = self.value(lua, i, j)?;
            if !v.is_nil() {
                t.raw_set(self.key_value(lua, j)?, v)?;
            }
        }
        Ok(t)
    }

    fn region_string(&self) -> Option<String> {
        (self.r2 >= self.r1 && self.c2 >= self.c1).then(|| {
            format!(
                "{}:{}",
                refs::a1(self.r1, self.c1),
                refs::a1(self.r2, self.c2)
            )
        })
    }

    /// New view over `region` sharing this grid.
    fn view(&self, region: (u32, u32, u32, u32), spec: HeaderSpec) -> SheetInner {
        SheetInner::new(self.grid.clone(), region, spec)
    }

    fn full_region(&self) -> (u32, u32, u32, u32) {
        (self.r1, self.c1, self.r2, self.c2)
    }
}

#[derive(Clone)]
pub struct Sheet(pub Rc<SheetInner>);

pub struct Row {
    sheet: Rc<SheetInner>,
    i: u32,
}

fn sheet_ud(lua: &Lua, s: SheetInner) -> Result<Value> {
    Ok(Value::UserData(lua.create_userdata(Sheet(Rc::new(s)))?))
}

fn row_ud(lua: &Lua, s: &Rc<SheetInner>, i: u32) -> Result<Value> {
    Ok(Value::UserData(lua.create_userdata(Row {
        sheet: s.clone(),
        i,
    })?))
}

/// Integer-valued Lua number -> index.
#[allow(clippy::unnecessary_cast)] // mlua::Integer is not i64 on every Lua flavour
fn as_index(v: &Value) -> Option<i64> {
    match v {
        Value::Integer(i) => Some(*i as i64),
        Value::Number(n) if n.fract() == 0.0 && n.abs() < 9.0e15 => Some(*n as i64),
        _ => None,
    }
}

fn pos_from(args: &[Value], what: &str) -> Result<Pos> {
    match args {
        [Value::String(s), ..] => {
            let s = s.to_string_lossy();
            refs::parse_cell(&s).ok_or_else(|| err(format!("invalid {what} reference '{s}'")))
        }
        [r, c, ..] => match (as_index(r), as_index(c)) {
            (Some(r), Some(c))
                if r >= 1 && c >= 1 && r <= refs::MAX_ROW as i64 && c <= refs::MAX_COL as i64 =>
            {
                Ok(Pos {
                    row: r as u32,
                    col: c as u32,
                })
            }
            _ => Err(err(format!("invalid {what} row/column"))),
        },
        _ => Err(err(format!(
            "expected a {what} reference (\"B7\", \"R7C2\") or row, col"
        ))),
    }
}

/// Parse "A1:C5", "R1C1:R5C3" or the open-ended "A5:" (to the end of the used range).
fn region_from_str(s: &str, grid: &Grid) -> Result<(u32, u32, u32, u32)> {
    let s = s.trim();
    if let Some(start) = s.strip_suffix(':') {
        let p = refs::parse_cell(start).ok_or_else(|| err(format!("invalid range '{s}'")))?;
        let (er, ec) = grid.used_end();
        return Ok((p.row, p.col, er.max(p.row), ec.max(p.col)));
    }
    let (a, b) = refs::parse_range(s).ok_or_else(|| err(format!("invalid range '{s}'")))?;
    Ok((a.row, a.col, b.row, b.col))
}

/// Region and header spec for a freshly opened sheet.
pub fn open_view(grid: Rc<Grid>, opts: Option<&Table>) -> Result<SheetInner> {
    let mut region = {
        let (er, ec) = grid.used_end();
        (1, 1, er, ec)
    };
    let mut spec = HeaderSpec::None;
    if let Some(o) = opts {
        if let Some(r) = o.get::<Option<String>>("range")? {
            region = region_from_str(&r, &grid)?;
        }
        spec = HeaderSpec::from_value(o.get::<Value>("headers")?)?;
    }
    Ok(SheetInner::new(grid, region, spec))
}

pub fn table_view(grid: Rc<Grid>, data: &Range<Data>, columns: &[String]) -> SheetInner {
    let region = match (data.start(), data.end()) {
        (Some(s), Some(e)) => (s.0 + 1, s.1 + 1, e.0 + 1, e.1 + 1),
        _ => (1, 1, 0, 0),
    };
    SheetInner::new(grid, region, HeaderSpec::Names(columns.to_vec()))
}

type Matcher = Box<dyn Fn(&Data, Fmt) -> Result<bool>>;

fn matcher(lua: &Lua, needle: Value, opts: Option<Table>) -> Result<Matcher> {
    let flag = |k: &str| -> Result<bool> {
        Ok(match &opts {
            Some(o) => o.get::<Option<bool>>(k)?.unwrap_or(false),
            None => false,
        })
    };
    match needle {
        Value::String(s) => {
            let text = s.to_string_lossy().to_string();
            if flag("pattern")? {
                let find: Function = lua.globals().get::<Table>("string")?.get("find")?;
                let lua = lua.clone();
                return Ok(Box::new(move |d, fmt| {
                    if let Data::String(cell) = d {
                        let s = lua.create_string(cell)?;
                        let hit: Value = find.call((s, text.as_str()))?;
                        Ok(!hit.is_nil())
                    } else {
                        let _ = fmt;
                        Ok(false)
                    }
                }));
            }
            let (contains, icase) = (flag("contains")?, flag("ignorecase")?);
            let text = if icase { text.to_lowercase() } else { text };
            Ok(Box::new(move |d, fmt| {
                if matches!(
                    d,
                    Data::Empty | Data::Int(_) | Data::Float(_) | Data::Bool(_)
                ) {
                    return Ok(false);
                }
                let mut cell = display(d, fmt);
                if icase {
                    cell = cell.to_lowercase();
                }
                Ok(if contains {
                    cell.contains(&text)
                } else {
                    cell == text
                })
            }))
        }
        Value::Integer(_) | Value::Number(_) => {
            let n = match needle {
                Value::Integer(i) => i as f64,
                Value::Number(n) => n,
                _ => unreachable!(),
            };
            Ok(Box::new(move |d, _| {
                Ok(match d {
                    Data::Int(i) => *i as f64 == n,
                    Data::Float(x) => *x == n,
                    _ => false,
                })
            }))
        }
        Value::Boolean(b) => Ok(Box::new(move |d, _| {
            Ok(matches!(d, Data::Bool(x) if *x == b))
        })),
        other => Err(err(format!("cannot search for a {}", other.type_name()))),
    }
}

impl Sheet {
    fn scan(
        &self,
        lua: &Lua,
        needle: Value,
        opts: Option<Table>,
        first_only: bool,
    ) -> Result<Vec<(u32, u32)>> {
        let s = &self.0;
        let hit = matcher(lua, needle, opts)?;
        let mut found = Vec::new();
        for i in 1..=s.nrows() {
            for j in 1..=s.ncols() {
                if let Some(d) = s.data(i, j) {
                    if hit(d, s.grid.fmt)? {
                        found.push((s.row0() + i - 1, s.c1 + j - 1));
                        if first_only {
                            return Ok(found);
                        }
                    }
                }
            }
        }
        Ok(found)
    }
}

impl UserData for Sheet {
    fn add_fields<F: UserDataFields<Self>>(fields: &mut F) {
        fields.add_field_method_get("name", |_, t| Ok(t.0.grid.name.clone()));
        fields.add_field_method_get("nrows", |_, t| Ok(t.0.nrows()));
        fields.add_field_method_get("ncols", |_, t| Ok(t.0.ncols()));
        fields.add_field_method_get("headers", |_, t| {
            Ok(t.0.headers.as_ref().map(|h| h.names.clone()))
        });
        fields.add_field_method_get("region", |_, t| Ok(t.0.region_string()));
        fields.add_field_method_get("dimension", |_, t| {
            let (er, ec) = t.0.grid.used_end();
            Ok((er > 0).then(|| {
                let (sr, sc) =
                    t.0.grid
                        .range
                        .start()
                        .map_or((0, 0), |s| (s.0 + 1, s.1 + 1));
                format!("{}:{}", refs::a1(sr, sc), refs::a1(er, ec))
            }))
        });
    }

    fn add_methods<M: UserDataMethods<Self>>(m: &mut M) {
        // sheet:cell("B7") / sheet:cell("R7C2") / sheet:cell(7, 2) -- absolute
        m.add_method("cell", |lua, t, args: MultiValue| {
            let args: Vec<Value> = args.into_iter().collect();
            let p = pos_from(&args, "cell")?;
            match t.0.grid.cell(p.row, p.col) {
                Some(d) => to_value(lua, d, t.0.grid.fmt),
                None => Ok(Value::Nil),
            }
        });

        // sheet:range("B3:F20" [, opts]) / sheet:range(r1, c1, r2, c2 [, opts]) -- absolute
        m.add_method("range", |lua, t, args: MultiValue| {
            let args: Vec<Value> = args.into_iter().collect();
            let (region, rest) = match args.first() {
                Some(Value::String(s)) => (
                    region_from_str(&s.to_string_lossy(), &t.0.grid)?,
                    &args[1..],
                ),
                Some(_) if args.len() >= 4 => {
                    let a = pos_from(&args[0..2], "range")?;
                    let b = pos_from(&args[2..4], "range")?;
                    (
                        (
                            a.row.min(b.row),
                            a.col.min(b.col),
                            a.row.max(b.row),
                            a.col.max(b.col),
                        ),
                        &args[4..],
                    )
                }
                _ => return Err(err("range expects \"A1:C5\" or row1, col1, row2, col2")),
            };
            let spec = match rest.first() {
                Some(Value::Table(o)) => HeaderSpec::from_value(o.get::<Value>("headers")?)?,
                _ => HeaderSpec::None,
            };
            sheet_ud(lua, t.0.view(region, spec))
        });

        // sheet:with_headers(true | false | {"a", "b"}) -- same region, new headers
        m.add_method("with_headers", |lua, t, spec: Value| {
            sheet_ud(
                lua,
                t.0.view(t.0.full_region(), HeaderSpec::from_value(spec)?),
            )
        });

        m.add_method("row", |lua, t, i: Value| match as_index(&i) {
            Some(i) if i >= 1 && i <= t.0.nrows() as i64 => row_ud(lua, &t.0, i as u32),
            _ => Ok(Value::Nil),
        });

        // sheet:col(2) / sheet:col("name") -> array of values down the window
        m.add_method("col", |lua, t, key: Value| {
            let s = &t.0;
            let j = match &key {
                Value::String(k) => {
                    let k = k.to_string_lossy();
                    s.headers
                        .as_ref()
                        .and_then(|h| h.index.get(k.as_str()).copied())
                        .ok_or_else(|| err(format!("no column with header '{k}'")))?
                }
                other => match as_index(other) {
                    Some(j) if j >= 1 && j <= s.ncols() as i64 => j as u32,
                    _ => {
                        return Err(err(
                            "column must be an index within the sheet or a header name",
                        ))
                    }
                },
            };
            let out = lua.create_table()?;
            for i in 1..=s.nrows() {
                let v = s.value(lua, i, j)?;
                if !v.is_nil() {
                    out.raw_set(i, v)?;
                }
            }
            Ok(out)
        });

        m.add_method("rows", |lua, t, ()| {
            let s = t.0.clone();
            let mut i = 0u32;
            lua.create_function_mut(move |lua, _: MultiValue| {
                i += 1;
                if i > s.nrows() {
                    Ok(Value::Nil)
                } else {
                    row_ud(lua, &s, i)
                }
            })
        });

        // Plain Lua tables: an array of rows; rows are keyed by header when headers are set.
        m.add_method("totable", |lua, t, ()| {
            let out = lua.create_table_with_capacity(t.0.nrows() as usize, 0)?;
            for i in 1..=t.0.nrows() {
                out.raw_set(i, t.0.row_table(lua, i)?)?;
            }
            Ok(out)
        });

        // sheet:find(value [, {contains, ignorecase, pattern}]) -> row, col (absolute) | nil
        m.add_method("find", |lua, t, (needle, opts): (Value, Option<Table>)| {
            Ok(match t.scan(lua, needle, opts, true)?.first() {
                Some(&(r, c)) => {
                    MultiValue::from_iter([Value::Integer(r as _), Value::Integer(c as _)])
                }
                None => MultiValue::from_iter([Value::Nil]),
            })
        });

        // sheet:findall(value [, opts]) -> { {row=, col=, ref=, value=}, ... }
        m.add_method(
            "findall",
            |lua, t, (needle, opts): (Value, Option<Table>)| {
                let out = lua.create_table()?;
                for (n, (r, c)) in t.scan(lua, needle, opts, false)?.into_iter().enumerate() {
                    let hit = lua.create_table()?;
                    hit.set("row", r)?;
                    hit.set("col", c)?;
                    hit.set("ref", refs::a1(r, c))?;
                    if let Some(d) = t.0.grid.cell(r, c) {
                        hit.set("value", to_value(lua, d, t.0.grid.fmt)?)?;
                    }
                    out.raw_set(n + 1, hit)?;
                }
                Ok(out)
            },
        );

        m.add_method("formula", |_, t, args: MultiValue| {
            let args: Vec<Value> = args.into_iter().collect();
            let p = pos_from(&args, "cell")?;
            t.0.grid.formula(p.row, p.col)
        });

        m.add_method("merged", |_, t, ()| {
            Ok(t.0
                .grid
                .merged()?
                .iter()
                .map(|d| {
                    let (a, b) = ((d.start.0 + 1, d.start.1 + 1), (d.end.0 + 1, d.end.1 + 1));
                    format!("{}:{}", refs::a1(a.0, a.1), refs::a1(b.0, b.1))
                })
                .collect::<Vec<_>>())
        });

        // sheet[3] -> row, sheet["B7"] -> cell, sheet["B3:F20"] -> sub-sheet
        m.add_meta_method("__index", |lua, t, key: Value| match &key {
            Value::String(k) => {
                let k = k.to_string_lossy();
                if k.contains(':') {
                    return match refs::parse_range(&k) {
                        Some((a, b)) => sheet_ud(
                            lua,
                            t.0.view((a.row, a.col, b.row, b.col), HeaderSpec::None),
                        ),
                        None => Ok(Value::Nil),
                    };
                }
                match refs::parse_cell(&k) {
                    Some(p) => match t.0.grid.cell(p.row, p.col) {
                        Some(d) => to_value(lua, d, t.0.grid.fmt),
                        None => Ok(Value::Nil),
                    },
                    None => Ok(Value::Nil),
                }
            }
            other => match as_index(other) {
                Some(i) if i >= 1 && i <= t.0.nrows() as i64 => row_ud(lua, &t.0, i as u32),
                _ => Ok(Value::Nil),
            },
        });

        m.add_meta_method("__len", |_, t, ()| Ok(t.0.nrows()));

        m.add_meta_method("__pairs", |lua, t, ()| {
            let s = t.0.clone();
            let mut i = 0u32;
            let next = lua.create_function_mut(move |lua, _: MultiValue| {
                i += 1;
                if i > s.nrows() {
                    Ok((Value::Nil, Value::Nil))
                } else {
                    Ok((Value::Integer(i as _), row_ud(lua, &s, i)?))
                }
            })?;
            Ok((next, Value::Nil, Value::Nil))
        });

        m.add_meta_method("__tostring", |_, t, ()| {
            Ok(format!(
                "xlsx.sheet({:?}, {} rows x {} cols)",
                t.0.grid.name,
                t.0.nrows(),
                t.0.ncols()
            ))
        });
    }
}

impl UserData for Row {
    fn add_methods<M: UserDataMethods<Self>>(m: &mut M) {
        // row[3] -> third column; row.name -> column with that header
        m.add_meta_method("__index", |lua, t, key: Value| {
            let s = &t.sheet;
            match &key {
                Value::String(k) => {
                    let k = k.to_string_lossy();
                    match s.headers.as_ref().and_then(|h| h.index.get(k.as_str())) {
                        Some(&j) => s.value(lua, t.i, j),
                        None => Ok(Value::Nil),
                    }
                }
                other => match as_index(other) {
                    Some(j) if j >= 1 && j <= s.ncols() as i64 => s.value(lua, t.i, j as u32),
                    _ => Ok(Value::Nil),
                },
            }
        });

        m.add_meta_method("__len", |_, t, ()| Ok(t.sheet.ncols()));

        // Non-empty cells, keyed by header (or column number without headers).
        m.add_meta_method("__pairs", |lua, t, ()| {
            let s = t.sheet.clone();
            let i = t.i;
            let mut j = 0u32;
            let next = lua.create_function_mut(move |lua, _: MultiValue| {
                while j < s.ncols() {
                    j += 1;
                    let v = s.value(lua, i, j)?;
                    if !v.is_nil() {
                        return Ok((s.key_value(lua, j)?, v));
                    }
                }
                Ok((Value::Nil, Value::Nil))
            })?;
            Ok((next, Value::Nil, Value::Nil))
        });

        m.add_meta_method("__tostring", |_, t, ()| Ok(format!("xlsx.row({})", t.i)));
    }
}
