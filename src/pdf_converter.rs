use genpdf::Alignment;
use genpdf::Element as _;
use genpdf::{elements, fonts, style, Document};
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};
use std::fs;
use std::path::{Path, PathBuf};

pub fn find_font_family() -> Result<fonts::FontFamily<fonts::FontData>, String> {
    let candidates = [
        ("/usr/share/fonts/liberation-sans-fonts", "LiberationSans"),
        ("/usr/share/fonts/google-carlito-fonts", "Carlito"),
        ("/usr/share/fonts/truetype/dejavu", "DejaVuSans"),
        ("/usr/share/fonts/dejavu", "DejaVuSans"),
        ("/usr/share/fonts/truetype/liberation", "LiberationSans"),
        ("/usr/share/fonts/TTF", "DejaVuSans"),
        ("/Library/Fonts", "Arial"),
        ("/System/Library/Fonts", "Helvetica"),
        ("C:\\Windows\\Fonts", "arial"),
    ];

    for (dir, name) in &candidates {
        let regular_path = Path::new(dir).join(format!("{}-Regular.ttf", name));
        if regular_path.exists() {
            if let Ok(family) = fonts::from_files(dir, name, None) {
                return Ok(family);
            }
        }
    }

    Err("Не знайдено системного TTF шрифту (LiberationSans, Carlito, DejaVuSans). Встановіть шрифти або перевірте шлях.".to_string())
}

pub fn convert_markdown_to_pdf(
    md_text: &str,
    output_path: &Path,
    source_dir: Option<&Path>,
    verbose: bool,
) -> Result<crate::converter::ConvertStats, String> {
    let font_family = find_font_family()?;
    let mut doc = Document::new(font_family);
    doc.set_title("Звіт");
    doc.set_minimal_conformance();
    doc.set_line_spacing(1.25);

    let mut decorator = genpdf::SimplePageDecorator::new();
    decorator.set_margins(12);
    decorator.set_header(|page| {
        let mut layout = elements::LinearLayout::vertical();
        if page > 1 {
            layout.push(
                elements::Paragraph::new(format!("Сторінка {}", page))
                    .aligned(Alignment::Right)
                    .styled(style::Style::new().with_font_size(8)),
            );
            layout.push(elements::Break::new(0.5));
        }
        layout
    });
    doc.set_page_decorator(decorator);

    let mut stats = crate::converter::ConvertStats {
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

    let mut current_paragraph: Option<elements::Paragraph> = None;
    let mut current_heading_level: Option<HeadingLevel> = None;
    let mut bold = false;
    let mut italic = false;
    let mut in_code_block = false;
    let mut code_lang = String::new();
    let mut code_content = String::new();

    let mut in_table = false;
    let mut table_headers: Vec<String> = Vec::new();
    let mut table_rows: Vec<Vec<String>> = Vec::new();
    let mut current_row: Vec<String> = Vec::new();
    let mut current_cell_text = String::new();
    let mut is_header_cell = false;

    let mut in_image = false;
    let mut image_url = String::new();

    for event in parser {
        match event {
            Event::Start(Tag::Heading { level, .. }) => {
                stats.headings += 1;
                current_heading_level = Some(level);
                current_paragraph = Some(elements::Paragraph::default());
            }
            Event::End(TagEnd::Heading(_)) => {
                if let Some(p) = current_paragraph.take() {
                    let size = match current_heading_level.take().unwrap_or(HeadingLevel::H2) {
                        HeadingLevel::H1 => 18,
                        HeadingLevel::H2 => 14,
                        HeadingLevel::H3 => 12,
                        _ => 10,
                    };
                    doc.push(elements::Break::new(0.5));
                    doc.push(p.styled(style::Style::new().bold().with_font_size(size)));
                    doc.push(elements::Break::new(0.4));
                }
            }

            Event::Start(Tag::Paragraph) => {
                if !in_table && !in_image {
                    stats.paragraphs += 1;
                    current_paragraph = Some(elements::Paragraph::default());
                }
            }
            Event::End(TagEnd::Paragraph) => {
                if !in_table && !in_image {
                    if let Some(p) = current_paragraph.take() {
                        doc.push(p.styled(style::Style::new().with_font_size(10)));
                        doc.push(elements::Break::new(0.3));
                    }
                }
            }

            Event::Start(Tag::Strong) => {
                bold = true;
            }
            Event::End(TagEnd::Strong) => {
                bold = false;
            }

            Event::Start(Tag::Emphasis) => {
                italic = true;
            }
            Event::End(TagEnd::Emphasis) => {
                italic = false;
            }

            Event::Start(Tag::Image { dest_url, .. }) => {
                in_image = true;
                image_url = dest_url.to_string();
            }
            Event::End(TagEnd::Image) => {
                in_image = false;
                if !image_url.is_empty() {
                    let img_path = if let Some(dir) = source_dir {
                        dir.join(&image_url)
                    } else {
                        PathBuf::from(&image_url)
                    };

                    if img_path.exists() {
                        if let Ok(img) = elements::Image::from_path(&img_path) {
                            if verbose {
                                println!("  [v] Вбудовування зображення у PDF: {}", img_path.display());
                            }
                            doc.push(elements::Break::new(0.4));
                            doc.push(img.with_scale(genpdf::Scale::new(0.55, 0.55)).with_alignment(Alignment::Center));
                            doc.push(elements::Break::new(0.4));
                        }
                    } else if verbose {
                        eprintln!("  [!] Зображення не знайдено: {}", img_path.display());
                    }
                    image_url.clear();
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
                        let tmp_file = std::env::temp_dir().join(format!("mermaid_rust_{}.png", std::process::id()));
                        if fs::write(&tmp_file, &png_bytes).is_ok() {
                            if let Ok(img) = elements::Image::from_path(&tmp_file) {
                                doc.push(elements::Break::new(0.5));
                                doc.push(img.with_scale(genpdf::Scale::new(0.55, 0.55)).with_alignment(Alignment::Center));
                                doc.push(elements::Break::new(0.5));
                            }
                            let _ = fs::remove_file(&tmp_file);
                        }
                    }
                } else {
                    stats.code_blocks += 1;
                    let mut p = elements::Paragraph::default();
                    for line in code_content.lines() {
                        p.push_styled(line.to_string(), style::Style::new().with_font_size(8));
                        p.push("\n");
                    }
                    doc.push(elements::Break::new(0.3));
                    doc.push(p.padded(genpdf::Margins::all(2)).framed());
                    doc.push(elements::Break::new(0.3));
                }
            }

            Event::Start(Tag::Table(_)) => {
                stats.tables += 1;
                in_table = true;
                table_headers.clear();
                table_rows.clear();
            }
            Event::End(TagEnd::Table) => {
                in_table = false;
                if !table_headers.is_empty() || !table_rows.is_empty() {
                    let col_count = if !table_headers.is_empty() {
                        table_headers.len()
                    } else {
                        table_rows.first().map(|r| r.len()).unwrap_or(1)
                    };

                    let weights: Vec<usize> = vec![1; col_count];
                    let mut layout = elements::TableLayout::new(weights);
                    layout.set_cell_decorator(elements::FrameCellDecorator::new(true, true, false));

                    // Header row
                    if !table_headers.is_empty() {
                        let mut row = layout.row();
                        for h in &table_headers {
                            let p = elements::Paragraph::new(h.trim())
                                .styled(style::Style::new().bold().with_font_size(8))
                                .padded(1);
                            row.push_element(p);
                        }
                        let _ = row.push();
                    }

                    // Body rows
                    for r in &table_rows {
                        let mut row = layout.row();
                        for (i, cell) in r.iter().enumerate() {
                            if i < col_count {
                                let p = elements::Paragraph::new(cell.trim())
                                    .styled(style::Style::new().with_font_size(8))
                                    .padded(1);
                                row.push_element(p);
                            }
                        }
                        let _ = row.push();
                    }

                    doc.push(elements::Break::new(0.5));
                    doc.push(layout);
                    doc.push(elements::Break::new(0.5));
                }
            }
            Event::Start(Tag::TableHead) => {
                is_header_cell = true;
                current_row.clear();
            }
            Event::End(TagEnd::TableHead) => {
                is_header_cell = false;
                table_headers = current_row.clone();
                current_row.clear();
            }
            Event::Start(Tag::TableRow) => {
                current_row.clear();
            }
            Event::End(TagEnd::TableRow) => {
                if !is_header_cell {
                    table_rows.push(current_row.clone());
                }
                current_row.clear();
            }
            Event::Start(Tag::TableCell) => {
                current_cell_text.clear();
            }
            Event::End(TagEnd::TableCell) => {
                current_row.push(current_cell_text.clone());
                current_cell_text.clear();
            }

            Event::Text(txt) => {
                if in_code_block {
                    code_content.push_str(&txt);
                } else if in_table {
                    current_cell_text.push_str(&txt);
                } else if in_image {
                    // image description
                } else if let Some(ref mut p) = current_paragraph {
                    let mut st = style::Style::new();
                    if bold {
                        st = st.bold();
                    }
                    if italic {
                        st = st.italic();
                    }
                    p.push_styled(txt.to_string(), st);
                }
            }

            Event::Code(txt) => {
                if in_table {
                    current_cell_text.push_str(&txt);
                } else if let Some(ref mut p) = current_paragraph {
                    p.push_styled(txt.to_string(), style::Style::new().bold());
                }
            }

            Event::SoftBreak | Event::HardBreak => {
                if in_table {
                    current_cell_text.push(' ');
                } else if let Some(ref mut p) = current_paragraph {
                    p.push(" ");
                }
            }

            Event::Rule => {
                doc.push(elements::Break::new(0.5));
            }

            _ => {}
        }
    }

    doc.render_to_file(output_path)
        .map_err(|e| format!("Помилка запису PDF файлу {}: {}", output_path.display(), e))?;

    Ok(stats)
}
