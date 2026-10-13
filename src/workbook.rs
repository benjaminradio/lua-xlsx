//! Workbook userdata and the module entry points.

use std::cell::{Cell, RefCell};
use std::io::Cursor;
use std::rc::Rc;

use calamine::{Reader, SheetVisible, Xlsx};
use mlua::{Lua, Result, Table, UserData, UserDataMethods, Value};

use crate::convert::Fmt;
use crate::sheet::{open_view, table_view, Grid, Sheet, SheetInner, Wb};

fn err<S: std::fmt::Display>(msg: S) -> mlua::Error {
    mlua::Error::runtime(msg)
}

pub struct Workbook(Rc<Wb>);

pub fn from_bytes(bytes: Vec<u8>, opts: Option<&Table>) -> Result<Workbook> {
    let fmt = Fmt::default().with_opts(opts)?;
    let fill_merged = match opts {
        Some(o) => o.get::<Option<bool>>("fill_merged")?.unwrap_or(false),
        None => false,
    };
    let xlsx =
        Xlsx::new(Cursor::new(bytes)).map_err(|e| err(format!("cannot read workbook: {e}")))?;
    Ok(Workbook(Rc::new(Wb {
        xlsx: RefCell::new(xlsx),
        fmt,
        fill_merged,
        tables_loaded: Cell::new(false),
    })))
}

impl Workbook {
    #[allow(clippy::unnecessary_cast)] // mlua::Integer is not i64 on every Lua flavour
    fn sheet_name(&self, which: &Value) -> Result<String> {
        let names = self.0.xlsx.borrow().sheet_names();
        match which {
            Value::String(s) => {
                let s = s.to_string_lossy();
                names
                    .into_iter()
                    .find(|n| n == s.as_str())
                    .ok_or_else(|| err(format!("no sheet named '{s}'")))
            }
            Value::Integer(_) | Value::Number(_) => {
                let i = match which {
                    Value::Integer(i) => *i as i64,
                    Value::Number(n) if n.fract() == 0.0 => *n as i64,
                    _ => return Err(err("sheet index must be an integer")),
                };
                usize::try_from(i - 1)
                    .ok()
                    .and_then(|i| names.get(i).cloned())
                    .ok_or_else(|| err(format!("sheet index {i} out of range")))
            }
            other => Err(err(format!(
                "sheet must be a name or index, got {}",
                other.type_name()
            ))),
        }
    }

    fn open_sheet(&self, lua: &Lua, which: Value, opts: Option<Table>) -> Result<Value> {
        let name = self.sheet_name(&which)?;
        let fmt = self.0.fmt.with_opts(opts.as_ref())?;
        let fill = match &opts {
            Some(o) => o
                .get::<Option<bool>>("fill_merged")?
                .unwrap_or(self.0.fill_merged),
            None => self.0.fill_merged,
        };
        let grid = Grid::load(&self.0, &name, fmt, fill)?;
        let view = open_view(grid, opts.as_ref())?;
        Ok(Value::UserData(lua.create_userdata(Sheet(Rc::new(view)))?))
    }

    fn load_tables(&self) -> Result<()> {
        if !self.0.tables_loaded.get() {
            self.0.xlsx.borrow_mut().load_tables().map_err(err)?;
            self.0.tables_loaded.set(true);
        }
        Ok(())
    }
}

impl UserData for Workbook {
    fn add_methods<M: UserDataMethods<Self>>(m: &mut M) {
        m.add_method("sheetnames", |_, t, ()| Ok(t.0.xlsx.borrow().sheet_names()));

        m.add_method("sheetinfo", |lua, t, ()| {
            let out = lua.create_table()?;
            for (i, s) in t.0.xlsx.borrow().sheets_metadata().iter().enumerate() {
                let e = lua.create_table()?;
                e.set("name", s.name.as_str())?;
                e.set(
                    "visibility",
                    match s.visible {
                        SheetVisible::Visible => "visible",
                        SheetVisible::Hidden => "hidden",
                        SheetVisible::VeryHidden => "veryhidden",
                    },
                )?;
                out.raw_set(i + 1, e)?;
            }
            Ok(out)
        });

        // wb:sheet("Name" | 1 [, {headers=, range=, dates=, errors=, fill_merged=}])
        m.add_method("sheet", |lua, t, (which, opts): (Value, Option<Table>)| {
            t.open_sheet(lua, which, opts)
        });

        m.add_method("tables", |lua, t, ()| {
            t.load_tables()?;
            let out = lua.create_table()?;
            let x = t.0.xlsx.borrow();
            for (i, name) in x.table_names().into_iter().enumerate() {
                out.raw_set(i + 1, name.as_str())?;
            }
            Ok(out)
        });

        // wb:table("Table1") -> sheet window over the table body, headers from the table
        m.add_method("table", |lua, t, (name, opts): (String, Option<Table>)| {
            t.load_tables()?;
            let table = t.0.xlsx.borrow_mut().table_by_name(&name).map_err(err)?;
            let fmt = t.0.fmt.with_opts(opts.as_ref())?;
            let fill = match &opts {
                Some(o) => o
                    .get::<Option<bool>>("fill_merged")?
                    .unwrap_or(t.0.fill_merged),
                None => t.0.fill_merged,
            };
            let grid = Grid::load(&t.0, table.sheet_name(), fmt, fill)?;
            let view: SheetInner = table_view(grid, table.data(), table.columns());
            Ok(Value::UserData(lua.create_userdata(Sheet(Rc::new(view)))?))
        });

        m.add_method("definednames", |lua, t, ()| {
            let out = lua.create_table()?;
            for (name, target) in t.0.xlsx.borrow().defined_names() {
                out.set(name.as_str(), target.as_str())?;
            }
            Ok(out)
        });

        // wb[1] / wb["Sheet1"] -> sheet with default options
        m.add_meta_method("__index", |lua, t, key: Value| match key {
            Value::String(_) | Value::Integer(_) | Value::Number(_) => match t.sheet_name(&key) {
                Ok(_) => t.open_sheet(lua, key, None),
                Err(_) => Ok(Value::Nil),
            },
            _ => Ok(Value::Nil),
        });

        m.add_meta_method("__len", |_, t, ()| {
            Ok(t.0.xlsx.borrow().sheet_names().len())
        });

        m.add_meta_method("__tostring", |_, t, ()| {
            Ok(format!(
                "xlsx.workbook({})",
                t.0.xlsx.borrow().sheet_names().join(", ")
            ))
        });
    }
}

pub fn open(lua: &Lua, (path, opts): (String, Option<Table>)) -> Result<Workbook> {
    let _ = lua;
    let bytes = std::fs::read(&path).map_err(|e| err(format!("cannot open '{path}': {e}")))?;
    from_bytes(bytes, opts.as_ref())
}

pub fn load(_: &Lua, (data, opts): (mlua::LuaString, Option<Table>)) -> Result<Workbook> {
    from_bytes(data.as_bytes().to_vec(), opts.as_ref())
}
