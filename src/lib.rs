//! `require "xlsx"` -- read Excel workbooks from Lua and Lus.

mod convert;
mod refs;
mod sheet;
mod workbook;

use mlua::{Lua, Result, Table};

use refs::{a1, col_name, col_number, parse_cell, r1c1};

fn err<S: std::fmt::Display>(msg: S) -> mlua::Error {
    mlua::Error::runtime(msg)
}

#[mlua::lua_module]
fn xlsx(lua: &Lua) -> Result<Table> {
    let m = lua.create_table()?;
    m.set("_VERSION", env!("CARGO_PKG_VERSION"))?;
    m.set("open", lua.create_function(workbook::open)?)?;
    m.set("load", lua.create_function(workbook::load)?)?;

    // xlsx.colname(28) -> "AB"
    m.set(
        "colname",
        lua.create_function(|_, n: i64| {
            u32::try_from(n)
                .ok()
                .filter(|n| (1..=refs::MAX_COL).contains(n))
                .map(col_name)
                .ok_or_else(|| err(format!("column {n} out of range")))
        })?,
    )?;

    // xlsx.colnumber("AB") -> 28
    m.set(
        "colnumber",
        lua.create_function(|_, s: String| {
            col_number(s.trim()).ok_or_else(|| err(format!("invalid column '{s}'")))
        })?,
    )?;

    // xlsx.address(row, col [, "r1c1"]) -> "B7" | "R7C2"
    m.set(
        "address",
        lua.create_function(|_, (row, col, style): (u32, u32, Option<String>)| {
            if !(1..=refs::MAX_ROW).contains(&row) || !(1..=refs::MAX_COL).contains(&col) {
                return Err(err("row or column out of range"));
            }
            match style.as_deref().map(str::to_ascii_lowercase).as_deref() {
                None | Some("a1") => Ok(a1(row, col)),
                Some("r1c1") => Ok(r1c1(row, col)),
                Some(other) => Err(err(format!(
                    "unknown style '{other}', expected \"a1\" or \"r1c1\""
                ))),
            }
        })?,
    )?;

    // xlsx.parseref("B7") or ("R7C2") -> 7, 2
    m.set(
        "parseref",
        lua.create_function(|_, s: String| {
            parse_cell(&s)
                .map(|p| (p.row, p.col))
                .ok_or_else(|| err(format!("invalid cell reference '{s}'")))
        })?,
    )?;
    Ok(m)
}
