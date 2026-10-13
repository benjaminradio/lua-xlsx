//! Cell reference parsing and formatting ("A1" and "R1C1" notations).

pub const MAX_ROW: u32 = 1_048_576;
pub const MAX_COL: u32 = 16_384;

/// A 1-based worksheet coordinate.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Pos {
    pub row: u32,
    pub col: u32,
}

/// Convert a 1-based column number to its letters (1 -> "A", 28 -> "AB").
pub fn col_name(mut col: u32) -> String {
    let mut out = Vec::new();
    while col > 0 {
        col -= 1;
        out.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

/// Convert column letters to a 1-based number ("AB" -> 28).
pub fn col_number(name: &str) -> Option<u32> {
    if name.is_empty() || name.len() > 3 {
        return None;
    }
    let mut n: u32 = 0;
    for b in name.bytes() {
        if !b.is_ascii_alphabetic() {
            return None;
        }
        n = n * 26 + (b.to_ascii_uppercase() - b'A' + 1) as u32;
    }
    (n <= MAX_COL).then_some(n)
}

fn parse_num(s: &str) -> Option<u32> {
    if s.is_empty() || s.len() > 7 || !s.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    s.parse().ok()
}

fn checked(row: u32, col: u32) -> Option<Pos> {
    ((1..=MAX_ROW).contains(&row) && (1..=MAX_COL).contains(&col)).then_some(Pos { row, col })
}

fn parse_a1(s: &str) -> Option<Pos> {
    let s = s.strip_prefix('$').unwrap_or(s);
    let split = s.find(|c: char| !c.is_ascii_alphabetic())?;
    let (letters, rest) = s.split_at(split);
    let rest = rest.strip_prefix('$').unwrap_or(rest);
    checked(parse_num(rest)?, col_number(letters)?)
}

fn parse_r1c1(s: &str) -> Option<Pos> {
    let s = s.strip_prefix(['R', 'r'])?;
    let c = s.find(['C', 'c'])?;
    checked(parse_num(&s[..c])?, parse_num(&s[c + 1..])?)
}

/// Parse a single cell reference in either A1 ("B7", "$B$7") or R1C1 ("R7C2") notation.
pub fn parse_cell(s: &str) -> Option<Pos> {
    let s = s.trim();
    parse_a1(s).or_else(|| parse_r1c1(s))
}

/// Parse "X" or "X:Y" into a normalised (top-left, bottom-right) pair.
pub fn parse_range(s: &str) -> Option<(Pos, Pos)> {
    let s = s.trim();
    let (a, b) = match s.split_once(':') {
        Some((a, b)) => (parse_cell(a)?, parse_cell(b)?),
        None => {
            let p = parse_cell(s)?;
            (p, p)
        }
    };
    Some((
        Pos {
            row: a.row.min(b.row),
            col: a.col.min(b.col),
        },
        Pos {
            row: a.row.max(b.row),
            col: a.col.max(b.col),
        },
    ))
}

pub fn a1(row: u32, col: u32) -> String {
    format!("{}{}", col_name(col), row)
}

pub fn r1c1(row: u32, col: u32) -> String {
    format!("R{row}C{col}")
}
