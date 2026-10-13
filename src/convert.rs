//! Conversion of calamine cell values into Lua values.

use calamine::{Data, ExcelDateTime};
use mlua::{Lua, Result, Table, Value};

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Dates {
    /// ISO 8601 strings: "2024-01-31", "2024-01-31T13:30:00", "13:30:00".
    Iso,
    /// Raw Excel serial numbers (days since 1899-12-30, or 1904-01-01).
    Serial,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Errors {
    /// Error cells become their text: "#DIV/0!", "#N/A", ...
    Text,
    /// Error cells become nil.
    Nil,
}

#[derive(Clone, Copy)]
pub struct Fmt {
    pub dates: Dates,
    pub errors: Errors,
}

impl Default for Fmt {
    fn default() -> Self {
        Fmt {
            dates: Dates::Iso,
            errors: Errors::Text,
        }
    }
}

impl Fmt {
    /// Overlay the `dates` / `errors` options found in an options table.
    pub fn with_opts(mut self, opts: Option<&Table>) -> Result<Fmt> {
        let Some(opts) = opts else { return Ok(self) };
        if let Some(d) = opts.get::<Option<String>>("dates")? {
            self.dates = match d.as_str() {
                "iso" => Dates::Iso,
                "serial" => Dates::Serial,
                other => {
                    return Err(mlua::Error::runtime(format!(
                        "bad option 'dates': expected \"iso\" or \"serial\", got \"{other}\""
                    )))
                }
            };
        }
        if let Some(e) = opts.get::<Option<String>>("errors")? {
            self.errors = match e.as_str() {
                "text" | "string" => Errors::Text,
                "nil" => Errors::Nil,
                other => {
                    return Err(mlua::Error::runtime(format!(
                        "bad option 'errors': expected \"text\" or \"nil\", got \"{other}\""
                    )))
                }
            };
        }
        Ok(self)
    }
}

fn iso(dt: &ExcelDateTime) -> String {
    if dt.is_duration() {
        let total = (dt.as_f64() * 86_400.0).round() as i64;
        let sign = if total < 0 { "-" } else { "" };
        let t = total.abs();
        return format!("{sign}{}:{:02}:{:02}", t / 3600, (t / 60) % 60, t % 60);
    }
    let (y, mo, d, h, mi, s, _ms) = dt.to_ymd_hms_milli();
    let serial = dt.as_f64();
    if (0.0..1.0).contains(&serial) {
        format!("{h:02}:{mi:02}:{s:02}")
    } else if h == 0 && mi == 0 && s == 0 {
        format!("{y:04}-{mo:02}-{d:02}")
    } else {
        format!("{y:04}-{mo:02}-{d:02}T{h:02}:{mi:02}:{s:02}")
    }
}

/// Text used for header cells and string comparisons.
pub fn display(d: &Data, fmt: Fmt) -> String {
    match d {
        Data::Empty => String::new(),
        Data::Int(i) => i.to_string(),
        Data::Float(x) if x.fract() == 0.0 && x.abs() < 1e15 => format!("{}", *x as i64),
        Data::Float(x) => x.to_string(),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => s.clone(),
        Data::Bool(b) => b.to_string(),
        Data::DateTime(dt) => match fmt.dates {
            Dates::Iso => iso(dt),
            Dates::Serial => dt.as_f64().to_string(),
        },
        Data::Error(e) => e.to_string(),
    }
}

/// Excel stores every number as a float; hand whole numbers to Lua as integers.
fn number(x: f64) -> Value {
    if x.fract() == 0.0 && x.abs() < 9_007_199_254_740_992.0 {
        Value::Integer(x as mlua::Integer)
    } else {
        Value::Number(x)
    }
}

pub fn to_value(lua: &Lua, d: &Data, fmt: Fmt) -> Result<Value> {
    Ok(match d {
        Data::Empty => Value::Nil,
        Data::Int(i) => Value::Integer(*i as mlua::Integer),
        Data::Float(x) => number(*x),
        Data::Bool(b) => Value::Boolean(*b),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => {
            Value::String(lua.create_string(s)?)
        }
        Data::DateTime(dt) => match fmt.dates {
            Dates::Serial => number(dt.as_f64()),
            Dates::Iso => Value::String(lua.create_string(iso(dt))?),
        },
        Data::Error(e) => match fmt.errors {
            Errors::Nil => Value::Nil,
            Errors::Text => Value::String(lua.create_string(e.to_string())?),
        },
    })
}
