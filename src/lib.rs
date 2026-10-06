//! Batch operations for python-docx.
//!
//! Python keeps owning the document model. Rust only turns large batches of
//! data into WordprocessingML and large XML into text, crossing the FFI
//! boundary O(1) times per operation.

use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyString, PyTuple};
use quick_xml::{Reader, XmlVersion, events::Event};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn push_escaped(s: &str, out: &mut String) {
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            c => out.push(c),
        }
    }
}

fn push_attr_escaped(s: &str, out: &mut String) {
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            c => out.push(c),
        }
    }
}

#[derive(Default)]
struct RunFmt {
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    all_caps: bool,
    small_caps: bool,
    vert_align: Option<&'static str>,
    highlight: Option<&'static str>,
    size_half_points: Option<u32>,
    color: Option<String>,
    font: Option<String>,
}

impl RunFmt {
    fn is_plain(&self) -> bool {
        !(self.bold
            || self.italic
            || self.underline
            || self.strike
            || self.all_caps
            || self.small_caps
            || self.vert_align.is_some()
            || self.highlight.is_some()
            || self.size_half_points.is_some()
            || self.color.is_some()
            || self.font.is_some())
    }
}

/// ST_HighlightColor values Word accepts (lowercased, underscores removed -> canonical).
const HIGHLIGHTS: &[(&str, &str)] = &[
    ("black", "black"),
    ("blue", "blue"),
    ("cyan", "cyan"),
    ("green", "green"),
    ("magenta", "magenta"),
    ("red", "red"),
    ("yellow", "yellow"),
    ("white", "white"),
    ("darkblue", "darkBlue"),
    ("darkcyan", "darkCyan"),
    ("darkgreen", "darkGreen"),
    ("darkmagenta", "darkMagenta"),
    ("darkred", "darkRed"),
    ("darkyellow", "darkYellow"),
    ("darkgray", "darkGray"),
    ("lightgray", "lightGray"),
];

/// A highlight is a color name (`"yellow"`, `"dark_blue"`) or anything with an `xml_value`
/// attribute, such as python-docx's `WD_COLOR_INDEX` members.
fn parse_highlight(v: &Bound<'_, PyAny>) -> PyResult<&'static str> {
    let raw: String = match v.extract::<String>() {
        Ok(s) => s,
        Err(_) => v.getattr("xml_value")?.extract()?,
    };
    let key = raw.replace('_', "").to_ascii_lowercase();
    HIGHLIGHTS
        .iter()
        .find(|(k, _)| *k == key)
        .map(|(_, canon)| *canon)
        .ok_or_else(|| PyValueError::new_err(format!("unknown highlight color {raw:?}")))
}

fn push_run(text: &str, bold: bool, italic: bool, out: &mut String) {
    push_run_fmt(
        text,
        &RunFmt {
            bold,
            italic,
            ..Default::default()
        },
        out,
    );
}

/// Children of `<w:rPr>` are emitted in schema order (the order python-docx uses):
/// rFonts, b, i, caps, smallCaps, strike, color, sz, highlight, u, vertAlign.
fn push_run_fmt(text: &str, f: &RunFmt, out: &mut String) {
    out.push_str("<w:r>");
    if !f.is_plain() {
        out.push_str("<w:rPr>");
        if let Some(font) = &f.font {
            out.push_str("<w:rFonts w:ascii=\"");
            push_attr_escaped(font, out);
            out.push_str("\" w:hAnsi=\"");
            push_attr_escaped(font, out);
            out.push_str("\"/>");
        }
        if f.bold {
            out.push_str("<w:b/>");
        }
        if f.italic {
            out.push_str("<w:i/>");
        }
        if f.all_caps {
            out.push_str("<w:caps/>");
        }
        if f.small_caps {
            out.push_str("<w:smallCaps/>");
        }
        if f.strike {
            out.push_str("<w:strike/>");
        }
        if let Some(c) = &f.color {
            out.push_str("<w:color w:val=\"");
            out.push_str(c);
            out.push_str("\"/>");
        }
        if let Some(sz) = f.size_half_points {
            out.push_str(&format!("<w:sz w:val=\"{sz}\"/>"));
        }
        if let Some(h) = f.highlight {
            out.push_str(&format!("<w:highlight w:val=\"{h}\"/>"));
        }
        if f.underline {
            out.push_str("<w:u w:val=\"single\"/>");
        }
        if let Some(v) = f.vert_align {
            out.push_str(&format!("<w:vertAlign w:val=\"{v}\"/>"));
        }
        out.push_str("</w:rPr>");
    }
    let preserve = text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace);
    out.push_str(if preserve {
        "<w:t xml:space=\"preserve\">"
    } else {
        "<w:t>"
    });
    push_escaped(text, out);
    out.push_str("</w:t></w:r>");
}

fn wrap(body: &str) -> String {
    format!("<w:body xmlns:w=\"{W_NS}\">{body}</w:body>")
}

fn flag(obj: &Bound<'_, PyAny>) -> PyResult<bool> {
    if obj.is_none() {
        Ok(false)
    } else {
        obj.is_truthy()
    }
}

fn type_err(what: &str, obj: &Bound<'_, PyAny>) -> PyErr {
    let ty = obj
        .get_type()
        .name()
        .map(|n| n.to_string())
        .unwrap_or_default();
    PyTypeError::new_err(format!(
        "{what} must be str, (text, bold, italic) or (text, {{format}}), got {ty}"
    ))
}

/// Parse `{"bold", "italic", "underline", "strike", "all_caps", "small_caps", "superscript",
/// "subscript", "highlight", "size" (pt), "color" ("RRGGBB"), "font"}`.
fn parse_fmt(d: &Bound<'_, PyDict>) -> PyResult<RunFmt> {
    let mut f = RunFmt::default();
    let (mut sup, mut sub) = (false, false);
    for (k, v) in d.iter() {
        let key: String = k.extract()?;
        match key.as_str() {
            "bold" => f.bold = flag(&v)?,
            "italic" => f.italic = flag(&v)?,
            "underline" => f.underline = flag(&v)?,
            "strike" => f.strike = flag(&v)?,
            "all_caps" => f.all_caps = flag(&v)?,
            "small_caps" => f.small_caps = flag(&v)?,
            "superscript" => sup = flag(&v)?,
            "subscript" => sub = flag(&v)?,
            "highlight" if !v.is_none() => f.highlight = Some(parse_highlight(&v)?),
            "highlight" => {}
            "size" if !v.is_none() => {
                let pt: f64 = v.extract()?;
                if !(pt > 0.0 && pt <= 1638.0) {
                    return Err(PyValueError::new_err("size must be in (0, 1638] points"));
                }
                f.size_half_points = Some((pt * 2.0) as u32);
            }
            "color" if !v.is_none() => {
                let s: String = v.extract()?;
                let s = s.trim_start_matches('#').to_ascii_uppercase();
                if s.len() != 6 || !s.bytes().all(|b| b.is_ascii_hexdigit()) {
                    return Err(PyValueError::new_err(format!(
                        "color must be 'RRGGBB', got {s:?}"
                    )));
                }
                f.color = Some(s);
            }
            "font" if !v.is_none() => f.font = Some(v.extract()?),
            "size" | "color" | "font" => {}
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown run format key {other:?}"
                )));
            }
        }
    }
    if sup && sub {
        return Err(PyValueError::new_err(
            "a run cannot be both superscript and subscript",
        ));
    }
    f.vert_align = if sup {
        Some("superscript")
    } else if sub {
        Some("subscript")
    } else {
        None
    };
    Ok(f)
}

/// run: `"text"`, `(text, bold, italic)` (flags may be None) or `(text, {format})`.
fn push_run_obj(run: &Bound<'_, PyAny>, out: &mut String) -> PyResult<()> {
    if let Ok(s) = run.cast::<PyString>() {
        push_run(&s.to_cow()?, false, false, out);
        return Ok(());
    }
    if let Ok(t) = run.cast::<PyTuple>() {
        let n = t.len();
        if n == 2 || n == 3 {
            let text = t.get_item(0)?;
            let text = text
                .cast::<PyString>()
                .map_err(|_| type_err("run text", &text))?;
            if n == 3 {
                let (b, i) = (flag(&t.get_item(1)?)?, flag(&t.get_item(2)?)?);
                push_run(&text.to_cow()?, b, i, out);
            } else {
                let fmt = t.get_item(1)?;
                let fmt = fmt
                    .cast::<PyDict>()
                    .map_err(|_| type_err("run format", &fmt))?;
                push_run_fmt(&text.to_cow()?, &parse_fmt(fmt)?, out);
            }
            return Ok(());
        }
    }
    Err(type_err("run", run))
}

/// `paras`: any iterable of paragraphs; a paragraph is `"text"` or an iterable of runs.
#[pyfunction]
#[pyo3(signature = (paras, style_id=None, alignment=None))]
fn paragraphs_xml(
    paras: &Bound<'_, PyAny>,
    style_id: Option<String>,
    alignment: Option<String>,
) -> PyResult<String> {
    let mut out = String::with_capacity(paras.len().unwrap_or(1024) * 200);
    for para in paras.try_iter()? {
        let para = para?;
        out.push_str("<w:p>");
        if style_id.is_some() || alignment.is_some() {
            // pPr children are in schema order: pStyle, jc
            out.push_str("<w:pPr>");
            if let Some(id) = &style_id {
                out.push_str("<w:pStyle w:val=\"");
                push_attr_escaped(id, &mut out);
                out.push_str("\"/>");
            }
            if let Some(jc) = &alignment {
                out.push_str("<w:jc w:val=\"");
                push_attr_escaped(jc, &mut out);
                out.push_str("\"/>");
            }
            out.push_str("</w:pPr>");
        }
        if let Ok(s) = para.cast::<PyString>() {
            push_run(&s.to_cow()?, false, false, &mut out);
        } else if let Ok(runs) = para.try_iter() {
            for run in runs {
                push_run_obj(&run?, &mut out)?;
            }
        } else {
            return Err(type_err("paragraph", &para));
        }
        out.push_str("</w:p>");
    }
    Ok(wrap(&out))
}

/// `rows`: list of rows of cells (any object; non-str cells are `str()`-ed).
/// `col_widths`: one grid/cell width in twips per column (python-docx's `add_table` splits the
/// text width evenly). `header`: render the first row in bold.
#[pyfunction]
#[pyo3(signature = (rows, col_widths, header=false))]
fn table_xml(
    rows: Vec<Vec<Bound<'_, PyAny>>>,
    col_widths: Vec<i64>,
    header: bool,
) -> PyResult<String> {
    let cols = rows.first().map_or(0, |r| r.len());
    if rows.iter().any(|r| r.len() != cols) {
        return Err(PyValueError::new_err(
            "all rows must have the same number of cells",
        ));
    }
    if col_widths.len() != cols {
        return Err(PyValueError::new_err(format!(
            "col_widths has {} entries but the table has {cols} columns",
            col_widths.len()
        )));
    }
    let mut out = String::with_capacity(rows.len() * cols * 160);
    out.push_str(
        "<w:tbl><w:tblPr><w:tblW w:type=\"auto\" w:w=\"0\"/>\
         <w:tblLook w:firstColumn=\"1\" w:firstRow=\"1\" w:lastColumn=\"0\" w:lastRow=\"0\" \
         w:noHBand=\"0\" w:noVBand=\"1\" w:val=\"04A0\"/></w:tblPr><w:tblGrid>",
    );
    for w in &col_widths {
        out.push_str(&format!("<w:gridCol w:w=\"{w}\"/>"));
    }
    out.push_str("</w:tblGrid>");
    let tc_opens: Vec<String> = col_widths
        .iter()
        .map(|w| format!("<w:tc><w:tcPr><w:tcW w:type=\"dxa\" w:w=\"{w}\"/></w:tcPr><w:p>"))
        .collect();
    for (r, row) in rows.iter().enumerate() {
        out.push_str("<w:tr>");
        for (c, cell) in row.iter().enumerate() {
            out.push_str(&tc_opens[c]);
            let text = cell.str()?;
            let text = text.to_cow()?;
            if !text.is_empty() {
                push_run(&text, header && r == 0, false, &mut out);
            }
            out.push_str("</w:p></w:tc>");
        }
        out.push_str("</w:tr>");
    }
    out.push_str("</w:tbl>");
    Ok(wrap(&out))
}

/// Text of every `<w:p>` in document order (paragraphs inside tables included).
///
/// The parse is pure Rust over an immutable `bytes` buffer, so it runs with the GIL released;
/// only turning the result into Python strings needs the GIL again.
#[pyfunction]
fn extract_paragraph_texts(py: Python<'_>, xml: &[u8]) -> PyResult<Vec<String>> {
    py.detach(|| extract_texts(xml))
        .map_err(PyValueError::new_err)
}

fn extract_texts(xml: &[u8]) -> Result<Vec<String>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut buf = Vec::new();
    let mut paras = Vec::new();
    let mut cur = String::new();
    let mut in_t = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == "w:t" => in_t = true,
            Ok(Event::End(e)) => match e.name().as_ref() {
                "w:t" => in_t = false,
                "w:p" => paras.push(std::mem::take(&mut cur)),
                _ => {}
            },
            Ok(Event::Empty(e)) if e.name().as_ref() == "w:p" => paras.push(String::new()),
            Ok(Event::Text(t)) if in_t => cur.push_str(&t.xml_content(XmlVersion::Implicit1_0)),
            Ok(Event::GeneralRef(r)) if in_t => match r.as_ref() {
                "amp" => cur.push('&'),
                "lt" => cur.push('<'),
                "gt" => cur.push('>'),
                "quot" => cur.push('"'),
                "apos" => cur.push('\''),
                other => {
                    // numeric character reference: &#N; or &#xH;
                    if let Some(num) = other.strip_prefix('#') {
                        let code = match num.strip_prefix(['x', 'X']) {
                            Some(h) => u32::from_str_radix(h, 16),
                            None => num.parse::<u32>(),
                        };
                        if let Some(c) = code.ok().and_then(char::from_u32) {
                            cur.push(c);
                        }
                    }
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.to_string()),
            _ => {}
        }
        buf.clear();
    }
    Ok(paras)
}

#[pymodule]
fn _core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(paragraphs_xml, m)?)?;
    m.add_function(wrap_pyfunction!(table_xml, m)?)?;
    m.add_function(wrap_pyfunction!(extract_paragraph_texts, m)?)?;
    Ok(())
}
