use docx_rs::*;
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};
use std::fs;
use std::path::Path;

struct CalloutStyle {
    label: &'static str,
    icon: &'static str,
    border_color: &'static str,
    bg_color: &'static str,
    text_color: &'static str,
}

fn get_callout_style(alert_type: &str) -> CalloutStyle {
    match alert_type {
        "NOTE" => CalloutStyle {
            label: "ПРИМІТКА",
            icon: "ⓘ",
            border_color: "007AFF",
            bg_color: "F0F6FF",
            text_color: "007AFF",
        },
        "TIP" => CalloutStyle {
            label: "ПОРАДА",
            icon: "💡",
            border_color: "34C759",
            bg_color: "F0FAF4",
            text_color: "248A3D",
        },
        "IMPORTANT" => CalloutStyle {
            label: "ВАЖЛИВО",
            icon: "❗",
            border_color: "AF52DE",
            bg_color: "FAF5FF",
            text_color: "8936B2",
        },
        "WARNING" => CalloutStyle {
            label: "ЗАСТЕРЕЖЕННЯ",
            icon: "⚠️",
            border_color: "FF9500",
            bg_color: "FFFBF2",
            text_color: "D97706",
        },
        "CAUTION" => CalloutStyle {
            label: "НЕБЕЗПЕКА",
            icon: "🛑",
            border_color: "FF3B30",
            bg_color: "FFF2F2",
            text_color: "D72318",
        },
        _ => CalloutStyle {
            label: "ПРИМІТКА",
            icon: "ⓘ",
            border_color: "007AFF",
            bg_color: "F0F6FF",
            text_color: "007AFF",
        },
    }
}

#[derive(Debug)]
pub struct ConvertStats {
    pub headings: usize,
    pub paragraphs: usize,
    pub tables: usize,
    pub code_blocks: usize,
    pub mermaid_diagrams: usize,
    pub callouts: usize,
}

pub fn convert_markdown_to_docx(
    md_text: &str,
    output_path: &Path,
    _source_dir: Option<&Path>,
    verbose: bool,
) -> Result<ConvertStats, String> {
    let mut doc = Docx::new();

    let mut stats = ConvertStats {
        headings: 0,
        paragraphs: 0,
        tables: 0,
        code_blocks: 0,
        mermaid_diagrams: 0,
        callouts: 0,
    };

    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);
    options.insert(Options::ENABLE_MATH);

    let parser = Parser::new_ext(md_text, options);

    let mut current_paragraph: Option<Paragraph> = None;
    let mut bold = false;
    let mut italic = false;
    let mut strike = false;
    let mut in_code_block = false;
    let mut code_lang = String::new();
    let mut code_content = String::new();

    let mut in_blockquote = false;
    let mut blockquote_events: Vec<Event<'static>> = Vec::new();

    let mut table_rows: Vec<TableRow> = Vec::new();
    let mut current_row: Vec<TableCell> = Vec::new();
    let mut is_header_row = false;

    for event in parser {
        if in_blockquote {
            match &event {
                Event::End(TagEnd::BlockQuote(_)) => {
                    in_blockquote = false;
                    render_blockquote(
                        &blockquote_events,
                        &mut doc,
                        &mut stats,
                        verbose,
                    );
                    blockquote_events.clear();
                }
                _ => {
                    blockquote_events.push(event.into_static());
                }
            }
            continue;
        }

        match event {
            Event::Start(Tag::BlockQuote(_)) => {
                in_blockquote = true;
                blockquote_events.clear();
            }
            Event::Start(Tag::Heading { level, .. }) => {
                let lvl_num = match level {
                    HeadingLevel::H1 => 1,
                    HeadingLevel::H2 => 2,
                    HeadingLevel::H3 => 3,
                    HeadingLevel::H4 => 4,
                    HeadingLevel::H5 => 5,
                    HeadingLevel::H6 => 6,
                };
                stats.headings += 1;
                current_paragraph = Some(
                    Paragraph::new()
                        .style(&format!("Heading{}", lvl_num))
                        .line_spacing(LineSpacing::new().before(160).after(80)),
                );
            }
            Event::End(TagEnd::Heading(_)) => {
                if let Some(p) = current_paragraph.take() {
                    doc = doc.add_paragraph(p);
                }
            }
            Event::Start(Tag::Paragraph) => {
                stats.paragraphs += 1;
                current_paragraph = Some(
                    Paragraph::new()
                        .line_spacing(LineSpacing::new().before(40).after(100)),
                );
            }
            Event::End(TagEnd::Paragraph) => {
                if let Some(p) = current_paragraph.take() {
                    doc = doc.add_paragraph(p);
                }
            }
            Event::Start(Tag::CodeBlock(kind)) => {
                in_code_block = true;
                code_lang = match kind {
                    pulldown_cmark::CodeBlockKind::Fenced(lang) => lang.to_string(),
                    _ => String::new(),
                };
                code_content.clear();
            }
            Event::End(TagEnd::CodeBlock) => {
                in_code_block = false;
                let lang_clean = code_lang.trim().to_lowercase();
                if lang_clean == "mermaid" {
                    stats.mermaid_diagrams += 1;
                    if let Some(png_bytes) = crate::mermaid::render_mermaid(&code_content, verbose) {
                        let pic = Pic::new(&png_bytes).size(500 * 9525, 280 * 9525);
                        let p = Paragraph::new()
                            .align(AlignmentType::Center)
                            .line_spacing(LineSpacing::new().before(120).after(120))
                            .add_run(Run::new().add_image(pic));
                        doc = doc.add_paragraph(p);
                    } else {
                        doc = render_code_block_box(doc, &code_content, &code_lang);
                    }
                } else {
                    stats.code_blocks += 1;
                    doc = render_code_block_box(doc, &code_content, &code_lang);
                }
            }
            Event::Start(Tag::Table(_)) => {
                stats.tables += 1;
                table_rows.clear();
            }
            Event::End(TagEnd::Table) => {
                if !table_rows.is_empty() {
                    let tbl = Table::new(table_rows.clone())
                        .margins(TableCellMargins::new().margin(100, 150, 100, 150));
                    doc = doc.add_table(tbl);
                }
                table_rows.clear();
            }
            Event::Start(Tag::TableHead) => {
                is_header_row = true;
                current_row.clear();
            }
            Event::End(TagEnd::TableHead) => {
                is_header_row = false;
                if !current_row.is_empty() {
                    table_rows.push(TableRow::new(current_row.clone()));
                    current_row.clear();
                }
            }
            Event::Start(Tag::TableRow) => {
                current_row.clear();
            }
            Event::End(TagEnd::TableRow) => {
                if !current_row.is_empty() {
                    table_rows.push(TableRow::new(current_row.clone()));
                    current_row.clear();
                }
            }
            Event::Start(Tag::TableCell) => {
                current_paragraph = Some(Paragraph::new().line_spacing(LineSpacing::new().before(0).after(0)));
            }
            Event::End(TagEnd::TableCell) => {
                let p = current_paragraph.take().unwrap_or_default();
                let mut cell = TableCell::new().add_paragraph(p);
                if is_header_row {
                    cell = cell.shading(Shading::new().fill("E7E7E7"));
                }
                current_row.push(cell);
            }
            Event::Start(Tag::Item) => {
                current_paragraph = Some(
                    Paragraph::new()
                        .indent(Some(360), None, None, None)
                        .line_spacing(LineSpacing::new().before(20).after(40)),
                );
            }
            Event::End(TagEnd::Item) => {
                if let Some(p) = current_paragraph.take() {
                    doc = doc.add_paragraph(p);
                }
            }
            Event::Text(txt) => {
                if in_code_block {
                    code_content.push_str(&txt);
                } else if let Some(p) = current_paragraph.take() {
                    let mut r = Run::new().add_text(txt.as_ref());
                    if bold || is_header_row {
                        r = r.bold();
                    }
                    if italic {
                        r = r.italic();
                    }
                    if strike {
                        r = r.strike();
                    }
                    current_paragraph = Some(p.add_run(r));
                }
            }
            Event::Code(code) => {
                if let Some(p) = current_paragraph.take() {
                    let r = Run::new()
                        .add_text(code.as_ref())
                        .fonts(RunFonts::new().ascii("Consolas").hi_ansi("Consolas"))
                        .size(19)
                        .color("0F172A");
                    current_paragraph = Some(p.add_run(r));
                }
            }
            Event::InlineMath(math) | Event::DisplayMath(math) => {
                if let Some(p) = current_paragraph.take() {
                    let r = Run::new().add_text(&format!("[формула: {}]", math)).italic();
                    current_paragraph = Some(p.add_run(r));
                }
            }
            Event::Start(Tag::Emphasis) => italic = true,
            Event::End(TagEnd::Emphasis) => italic = false,
            Event::Start(Tag::Strong) => bold = true,
            Event::End(TagEnd::Strong) => bold = false,
            Event::Start(Tag::Strikethrough) => strike = true,
            Event::End(TagEnd::Strikethrough) => strike = false,
            Event::SoftBreak | Event::HardBreak => {
                if let Some(p) = current_paragraph.take() {
                    current_paragraph = Some(p.add_run(Run::new().add_break(BreakType::TextWrapping)));
                }
            }
            _ => {}
        }
    }

    if verbose {
        println!("  [v] Побудова DOCX дерева: {} абзаців, {} заголовків, {} таблиць", stats.paragraphs, stats.headings, stats.tables);
    }

    if let Some(parent) = output_path.parent() {
        let _ = fs::create_dir_all(parent);
    }
    let mut file = fs::File::create(output_path)
        .map_err(|e| format!("Не вдалося створити {}: {}", output_path.display(), e))?;

    doc.build()
        .pack(&mut file)
        .map_err(|e| format!("Помилка генерації DOCX: {}", e))?;

    if verbose {
        if let Ok(meta) = fs::metadata(output_path) {
            println!("  [v] Пакет DOCX збережено: {} (розмір: {} байт)", output_path.display(), meta.len());
        }
    }

    Ok(stats)
}

fn render_code_block_box(doc: Docx, code: &str, lang: &str) -> Docx {
    let mut cell = TableCell::new()
        .shading(Shading::new().fill("F8FAFC"))
        .set_border(
            TableCellBorder::new(TableCellBorderPosition::Left)
                .border_type(BorderType::Single)
                .size(4)
                .color("E2E8F0"),
        )
        .set_border(
            TableCellBorder::new(TableCellBorderPosition::Top)
                .border_type(BorderType::Single)
                .size(4)
                .color("E2E8F0"),
        )
        .set_border(
            TableCellBorder::new(TableCellBorderPosition::Right)
                .border_type(BorderType::Single)
                .size(4)
                .color("E2E8F0"),
        )
        .set_border(
            TableCellBorder::new(TableCellBorderPosition::Bottom)
                .border_type(BorderType::Single)
                .size(4)
                .color("E2E8F0"),
        );

    if !lang.trim().is_empty() {
        let p_lang = Paragraph::new()
            .align(AlignmentType::Right)
            .line_spacing(LineSpacing::new().before(0).after(40))
            .add_run(
                Run::new()
                    .add_text(lang.trim().to_uppercase())
                    .size(16)
                    .color("94A3B8"),
            );
        cell = cell.add_paragraph(p_lang);
    }

    for line in code.lines() {
        let line_txt = if line.is_empty() { " " } else { line };
        let p_line = Paragraph::new()
            .line_spacing(LineSpacing::new().before(0).after(0))
            .add_run(
                Run::new()
                    .add_text(line_txt)
                    .fonts(RunFonts::new().ascii("Consolas").hi_ansi("Consolas"))
                    .size(18)
                    .color("1E293B"),
            );
        cell = cell.add_paragraph(p_line);
    }

    let table = Table::new(vec![TableRow::new(vec![cell])])
        .margins(TableCellMargins::new().margin(120, 180, 120, 180));
    doc.add_table(table)
}

fn render_blockquote(
    events: &[Event<'static>],
    doc: &mut Docx,
    stats: &mut ConvertStats,
    verbose: bool,
) {
    let mut full_text = String::new();
    for ev in events {
        if let Event::Text(t) = ev {
            full_text.push_str(t);
        }
    }

    let alert_types = ["NOTE", "TIP", "IMPORTANT", "WARNING", "CAUTION"];
    let mut detected_type = None;
    for &at in &alert_types {
        let marker = format!("[!{}]", at);
        if full_text.trim_start().starts_with(&marker) {
            detected_type = Some(at);
            break;
        }
    }

    if let Some(at) = detected_type {
        stats.callouts += 1;
        let style = get_callout_style(at);
        if verbose {
            println!("  [v] Стилізація Callout-блоку [!{}] -> {}", at, style.label);
        }
        let cell = TableCell::new()
            .shading(Shading::new().fill(style.bg_color))
            .set_border(
                TableCellBorder::new(TableCellBorderPosition::Left)
                    .border_type(BorderType::Single)
                    .size(24)
                    .color(style.border_color),
            )
            .clear_border(TableCellBorderPosition::Top)
            .clear_border(TableCellBorderPosition::Right)
            .clear_border(TableCellBorderPosition::Bottom);

        let p_title = Paragraph::new()
            .line_spacing(LineSpacing::new().before(20).after(40))
            .add_run(
                Run::new()
                    .add_text(&format!("{}  {}", style.icon, style.label))
                    .bold()
                    .size(19)
                    .color(style.text_color),
            );

        let mut inner_cell = cell.add_paragraph(p_title);

        let marker = format!("[!{}]", at);
        let mut skipped_marker = false;
        let mut current_p = Paragraph::new().line_spacing(LineSpacing::new().before(20).after(40));

        for ev in events {
            match ev {
                Event::Text(t) => {
                    let mut text_to_add = t.as_ref();
                    if !skipped_marker && text_to_add.contains(&marker) {
                        text_to_add = text_to_add.replace(&marker, "").trim_start().to_string().leak();
                        skipped_marker = true;
                    }
                    if !text_to_add.trim().is_empty() {
                        current_p = current_p.add_run(Run::new().add_text(text_to_add));
                    }
                }
                Event::End(TagEnd::Paragraph) => {
                    inner_cell = inner_cell.add_paragraph(current_p);
                    current_p = Paragraph::new().line_spacing(LineSpacing::new().before(20).after(40));
                }
                _ => {}
            }
        }
        if !current_p.raw_text().is_empty() {
            inner_cell = inner_cell.add_paragraph(current_p);
        }

        let table = Table::new(vec![TableRow::new(vec![inner_cell])])
            .margins(TableCellMargins::new().margin(140, 200, 140, 200));
        *doc = doc.clone().add_table(table);
    } else {
        let mut p = Paragraph::new()
            .indent(Some(360), None, None, None)
            .line_spacing(LineSpacing::new().before(40).after(40));
        for ev in events {
            if let Event::Text(t) = ev {
                p = p.add_run(Run::new().add_text(t.as_ref()).italic());
            }
        }
        *doc = doc.clone().add_paragraph(p);
    }
}
