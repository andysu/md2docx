use genpdf::Alignment;
use genpdf::Element as _;
use genpdf::{elements, fonts, style, Document};
use image::GenericImageView;
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};
use std::path::{Path, PathBuf};

pub fn find_font_family() -> Result<fonts::FontFamily<fonts::FontData>, String> {
    // 1. Перший пріоритет: DejaVuSans (містить повну підтримку символів: ✔, ✕, ⚠ тощо)
    let dejavu_dirs = [
        "/usr/share/fonts/dejavu-sans-fonts",
        "/usr/share/fonts/dejavu",
        "/usr/share/fonts/truetype/dejavu",
    ];
    for dir in &dejavu_dirs {
        let p_reg = Path::new(dir).join("DejaVuSans.ttf");
        let p_bold = Path::new(dir).join("DejaVuSans-Bold.ttf");
        let p_italic = Path::new(dir).join("DejaVuSans-Oblique.ttf");
        let p_bold_italic = Path::new(dir).join("DejaVuSans-BoldOblique.ttf");
        if p_reg.exists() && p_bold.exists() {
            if let (Ok(regular), Ok(bold), Ok(italic), Ok(bold_italic)) = (
                fonts::FontData::load(&p_reg, None),
                fonts::FontData::load(&p_bold, None),
                fonts::FontData::load(if p_italic.exists() { &p_italic } else { &p_reg }, None),
                fonts::FontData::load(if p_bold_italic.exists() { &p_bold_italic } else { &p_bold }, None),
            ) {
                return Ok(fonts::FontFamily {
                    regular,
                    bold,
                    italic,
                    bold_italic,
                });
            }
        }
    }

    let candidates = [
        ("/usr/share/fonts/liberation-sans-fonts", "LiberationSans"),
        ("/usr/share/fonts/google-carlito-fonts", "Carlito"),
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

    Err("Не знайдено системного TTF шрифту (DejaVuSans, LiberationSans, Carlito). Встановіть шрифти або перевірте шлях.".to_string())
}

fn convert_dynamic_image_to_genpdf(dyn_img: image::DynamicImage) -> Result<elements::Image, String> {
    let rgb_img = if dyn_img.color().has_alpha() {
        let rgba = dyn_img.to_rgba8();
        let mut rgb = image::RgbImage::new(rgba.width(), rgba.height());
        for (x, y, pixel) in rgba.enumerate_pixels() {
            let alpha = pixel[3] as f32 / 255.0;
            let r = (pixel[0] as f32 * alpha + 255.0 * (1.0 - alpha)) as u8;
            let g = (pixel[1] as f32 * alpha + 255.0 * (1.0 - alpha)) as u8;
            let b = (pixel[2] as f32 * alpha + 255.0 * (1.0 - alpha)) as u8;
            rgb.put_pixel(x, y, image::Rgb([r, g, b]));
        }
        image::DynamicImage::ImageRgb8(rgb)
    } else {
        image::DynamicImage::ImageRgb8(dyn_img.to_rgb8())
    };

    let (w, _h) = (rgb_img.width() as f64, rgb_img.height() as f64);
    // Ширина сторінки A4 = 210 мм, поля по 12 мм з кожного боку => доступно 186 мм.
    // За формулою genpdf: width_mm = 25.4 * (scale * px_width) / 300.0.
    // Масштабуємо зображення точно на повну ширину сторінки (182 мм):
    let target_width_mm: f64 = 182.0;
    let scale: f64 = if w > 0.0 {
        (target_width_mm * 300.0) / (25.4 * w)
    } else {
        1.0
    };

    elements::Image::from_dynamic_image(rgb_img)
        .map(|img| img.with_scale(genpdf::Scale::new(scale, scale)).with_alignment(Alignment::Center))
        .map_err(|e| format!("Помилка створення Image для genpdf: {:?}", e))
}

fn load_image_rgb(path: &Path) -> Result<elements::Image, String> {
    let dyn_img = image::open(path).map_err(|e| format!("Помилка читання файлу {}: {}", path.display(), e))?;
    convert_dynamic_image_to_genpdf(dyn_img)
}

fn load_image_bytes_rgb(bytes: &[u8]) -> Result<elements::Image, String> {
    let dyn_img = image::load_from_memory(bytes).map_err(|e| format!("Помилка декодування зображення: {}", e))?;
    convert_dynamic_image_to_genpdf(dyn_img)
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

    let normalized_text = md_text
        .replace("✅", "✔ ")
        .replace("✔️", "✔ ")
        .replace("❌", "✕ ")
        .replace("⚠️", "⚠ ")
        .replace("\u{fe0f}", ""); // видалення variation selector

    let parser = Parser::new_ext(&normalized_text, options);

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
    let mut paragraph_has_text = false;

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
                    paragraph_has_text = false;
                    current_paragraph = Some(elements::Paragraph::default());
                }
            }
            Event::End(TagEnd::Paragraph) => {
                if !in_table && !in_image {
                    if let Some(p) = current_paragraph.take() {
                        if paragraph_has_text {
                            stats.paragraphs += 1;
                            doc.push(p.styled(style::Style::new().with_font_size(10)));
                            doc.push(elements::Break::new(0.3));
                        }
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
                        match load_image_rgb(&img_path) {
                            Ok(img) => {
                                if verbose {
                                    println!("  [v] Вбудовування зображення у PDF: {}", img_path.display());
                                }
                                doc.push(elements::Break::new(0.4));
                                doc.push(img);
                                doc.push(elements::Break::new(0.4));
                            }
                            Err(e) => {
                                eprintln!("  [!] Помилка завантаження зображення {}: {}", img_path.display(), e);
                            }
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
                        match load_image_bytes_rgb(&png_bytes) {
                            Ok(img) => {
                                doc.push(elements::Break::new(0.5));
                                doc.push(img);
                                doc.push(elements::Break::new(0.5));
                            }
                            Err(e) => {
                                if verbose {
                                    eprintln!("  [!] Помилка вбудовування діаграми Mermaid: {}", e);
                                }
                            }
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
                    if !txt.trim().is_empty() {
                        paragraph_has_text = true;
                    }
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
                    if !txt.trim().is_empty() {
                        paragraph_has_text = true;
                    }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_image_events() {
        let md = "# Тест\n\n![Щоденна динаміка](charts/daily_generation_timeline.png)\n";
        let out_pdf = std::env::temp_dir().join("test_img_render.pdf");
        let source_dir = Path::new("/home/andy/garb/analyse/teplomerezha");
        let res = convert_markdown_to_pdf(md, &out_pdf, Some(source_dir), true);
        assert!(res.is_ok(), "Не вдалося сконвертувати PDF з графіком: {:?}", res);
        assert!(out_pdf.exists());
        let _ = std::fs::remove_file(&out_pdf);
    }
}
