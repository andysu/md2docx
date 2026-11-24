use std::fs;
use std::path::{Path, PathBuf};
use pulldown_cmark::{Alignment, Event, HeadingLevel, Parser, Tag, TagEnd};
use typst::diag::{FileError, FileResult};
use typst::foundations::{Bytes, Datetime, Duration};
use typst::syntax::{FileId, RootedPath, Source, VirtualPath, VirtualRoot};
use typst::text::{Font, FontBook};
use typst::utils::LazyHash;
use typst::{Library, LibraryExt};
use typst::World;
use typst_layout::PagedDocument;
use typst_pdf::{pdf, PdfOptions};

pub struct TypstWorld {
    library: LazyHash<Library>,
    book: LazyHash<FontBook>,
    fonts: Vec<Font>,
    main_id: FileId,
    main_source: Source,
    source_dir: Option<PathBuf>,
}

impl TypstWorld {
    pub fn new(typst_text: String, source_dir: Option<PathBuf>) -> Result<Self, String> {
        let library = LazyHash::new(Library::default());

        let mut book = FontBook::new();
        let mut fonts = Vec::new();

        // 1. Вбудовані шрифти typst-assets
        for bytes in typst_assets::fonts() {
            let bytes = Bytes::new(bytes);
            for font in Font::iter(bytes) {
                book.push(font.info().clone());
                fonts.push(font);
            }
        }

        // 2. Системні шрифти (DejaVu Sans, Liberation Sans, тощо)
        let font_candidates = [
            "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans.ttf",
            "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Oblique.ttf",
            "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-BoldOblique.ttf",
            "/usr/share/fonts/liberation-sans-fonts/LiberationSans-Regular.ttf",
            "/usr/share/fonts/liberation-sans-fonts/LiberationSans-Bold.ttf",
            "/usr/share/fonts/liberation-sans-fonts/LiberationSans-Italic.ttf",
            "/usr/share/fonts/liberation-sans-fonts/LiberationSans-BoldItalic.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        ];

        for path_str in &font_candidates {
            let p = Path::new(path_str);
            if p.exists() {
                if let Ok(data) = fs::read(p) {
                    let bytes = Bytes::new(data);
                    for font in Font::iter(bytes) {
                        book.push(font.info().clone());
                        fonts.push(font);
                    }
                }
            }
        }

        let vpath = VirtualPath::new("main.typ").map_err(|e| format!("{:?}", e))?;
        let rooted = RootedPath::new(VirtualRoot::Project, vpath);
        let main_id = FileId::new(rooted);
        let main_source = Source::new(main_id, typst_text);

        Ok(Self {
            library,
            book: LazyHash::new(book),
            fonts,
            main_id,
            main_source,
            source_dir,
        })
    }
}

impl World for TypstWorld {
    fn library(&self) -> &LazyHash<Library> {
        &self.library
    }

    fn book(&self) -> &LazyHash<FontBook> {
        &self.book
    }

    fn main(&self) -> FileId {
        self.main_id
    }

    fn source(&self, id: FileId) -> FileResult<Source> {
        if id == self.main_id {
            Ok(self.main_source.clone())
        } else {
            Err(FileError::NotFound(PathBuf::from(id.vpath().get_without_slash())))
        }
    }

    fn file(&self, id: FileId) -> FileResult<Bytes> {
        if id == self.main_id {
            return Ok(Bytes::from_string(self.main_source.text().to_string()));
        }

        let rel_path_str = id.vpath().get_without_slash();
        let file_path = if let Some(ref base) = self.source_dir {
            base.join(rel_path_str)
        } else {
            PathBuf::from(rel_path_str)
        };

        if file_path.exists() {
            fs::read(&file_path)
                .map(Bytes::new)
                .map_err(|e| FileError::from_io(e, &file_path))
        } else {
            Err(FileError::NotFound(file_path))
        }
    }

    fn font(&self, index: usize) -> Option<Font> {
        self.fonts.get(index).cloned()
    }

    fn today(&self, _offset: Option<Duration>) -> Option<Datetime> {
        None
    }
}

/// Екранування спеціальних символів Typst у звичайному тексті
pub fn escape_typst_text(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '#' => out.push_str("\\#"),
            '[' => out.push_str("\\["),
            ']' => out.push_str("\\]"),
            '\\' => out.push_str("\\\\"),
            '*' => out.push_str("\\*"),
            '_' => out.push_str("\\_"),
            '@' => out.push_str("\\@"),
            '$' => out.push_str("\\$"),
            '<' => out.push_str("\\<"),
            '>' => out.push_str("\\>"),
            _ => out.push(ch),
        }
    }
    out
}

/// Заміна емодзі-статусів на чіткі векторні символи
pub fn normalize_status_icons(text: &str) -> String {
    text.replace("✅", "[✓]")
        .replace("☑", "[✓]")
        .replace("⚠️", "[!]")
        .replace("❌", "[✕]")
        .replace("✖", "[✕]")
        .replace("ℹ️", "[i]")
        .replace("ℹ", "[i]")
        .replace("$U_{mpp}$", "U_mpp")
        .replace("$I_{mpp}$", "I_mpp")
}

#[derive(Default)]
pub struct TypstStats {
    pub images: usize,
    pub tables: usize,
    pub callouts: usize,
}

/// Перетворює Markdown в Typst документ
pub fn markdown_to_typst(md_text: &str) -> (String, TypstStats) {
    let normalized = normalize_status_icons(md_text);
    let mut stats = TypstStats::default();

    let mut out = String::new();

    // Преамбула Typst: налаштування сторінки, шрифтів, відступів та стилів
    out.push_str(r##"
#set page(
  paper: "a4",
  margin: (x: 18mm, top: 20mm, bottom: 20mm),
  header: align(right)[
    #text(size: 8pt, fill: rgb("#718096"), font: ("DejaVu Sans", "Liberation Sans"))[
      Аудит СЕС «Тепломережа» | Конфіденційно
    ]
  ],
  footer: context {
    let page_number = counter(page).get().first()
    let total_pages = counter(page).final().first()
    set text(size: 8.5pt, fill: rgb("#718096"), font: ("DejaVu Sans", "Liberation Sans"))
    grid(
      columns: (1fr, auto),
      align(left)[Звіт інженерного аналізу генерації],
      align(right)[Сторінка #page_number з #total_pages]
    )
  }
)

#set text(
  font: ("DejaVu Sans", "Liberation Sans"),
  size: 9.5pt,
  fill: rgb("#2d3748"),
  lang: "uk"
)

#set par(justify: true, leading: 0.65em)

#show heading: set text(fill: rgb("#1a202c"), font: ("DejaVu Sans", "Liberation Sans"))
#show heading.where(level: 1): it => block(spacing: 1.5em)[
  #text(size: 15pt, weight: "bold", fill: rgb("#0d6efd"))[#it.body]
  #v(0.3em)
  #line(length: 100%, stroke: 0.75pt + rgb("#0d6efd"))
]
#show heading.where(level: 2): it => block(spacing: 1.2em)[
  #text(size: 12pt, weight: "bold", fill: rgb("#1a202c"))[#it.body]
]
#show heading.where(level: 3): it => block(spacing: 1.0em)[
  #text(size: 10.5pt, weight: "bold", fill: rgb("#2d3748"))[#it.body]
]

"##);

    let parser = Parser::new(&normalized);

    let mut in_table = false;
    let mut table_cells: Vec<Vec<String>> = Vec::new();
    let mut current_row: Vec<String> = Vec::new();
    let mut current_cell = String::new();
    let mut _in_table_head = false;
    let mut _table_alignments: Vec<Alignment> = Vec::new();

    let mut in_blockquote = false;
    let mut blockquote_content = String::new();

    let mut in_code_block = false;
    let mut code_block_content = String::new();
    let mut code_lang = String::new();

    struct ListState {
        is_ordered: bool,
        current_num: u64,
    }
    let mut list_stack: Vec<ListState> = Vec::new();

    for event in parser {
        match event {
            Event::Start(tag) => match tag {
                Tag::Heading { level, .. } => {
                    let prefix = match level {
                        HeadingLevel::H1 => "= ",
                        HeadingLevel::H2 => "== ",
                        HeadingLevel::H3 => "=== ",
                        HeadingLevel::H4 => "==== ",
                        HeadingLevel::H5 => "===== ",
                        HeadingLevel::H6 => "====== ",
                    };
                    out.push_str(prefix);
                }
                Tag::Paragraph => {}
                Tag::BlockQuote(_) => {
                    in_blockquote = true;
                    blockquote_content.clear();
                }
                Tag::CodeBlock(kind) => {
                    in_code_block = true;
                    code_block_content.clear();
                    code_lang = match kind {
                        pulldown_cmark::CodeBlockKind::Fenced(lang) => lang.to_string(),
                        pulldown_cmark::CodeBlockKind::Indented => String::new(),
                    };
                }
                Tag::List(first_num) => {
                    let is_ordered = first_num.is_some();
                    let current_num = first_num.unwrap_or(1);
                    list_stack.push(ListState { is_ordered, current_num });
                }
                Tag::Item => {
                    let indent = "  ".repeat(list_stack.len().saturating_sub(1));
                    let prefix = if let Some(state) = list_stack.last_mut() {
                        if state.is_ordered {
                            let n = state.current_num;
                            state.current_num += 1;
                            format!("{}. ", n)
                        } else {
                            "- ".to_string()
                        }
                    } else {
                        "- ".to_string()
                    };
                    out.push_str(&format!("{}{}", indent, prefix));
                }
                Tag::Table(aligns) => {
                    in_table = true;
                    table_cells.clear();
                    current_row.clear();
                    _table_alignments = aligns;
                    stats.tables += 1;
                }
                Tag::TableHead => {
                    _in_table_head = true;
                }
                Tag::TableRow => {
                    current_row.clear();
                }
                Tag::TableCell => {
                    current_cell.clear();
                }
                Tag::Emphasis => {
                    if in_blockquote {
                        blockquote_content.push('_');
                    } else if in_table {
                        current_cell.push('_');
                    } else {
                        out.push('_');
                    }
                }
                Tag::Strong => {
                    if in_blockquote {
                        blockquote_content.push('*');
                    } else if in_table {
                        current_cell.push('*');
                    } else {
                        out.push('*');
                    }
                }
                Tag::Strikethrough => {
                    if in_blockquote {
                        blockquote_content.push_str("#strike[");
                    } else if in_table {
                        current_cell.push_str("#strike[");
                    } else {
                        out.push_str("#strike[");
                    }
                }
                Tag::Link { dest_url, .. } => {
                    let s = format!("#link(\"{}\")[", dest_url);
                    if in_blockquote {
                        blockquote_content.push_str(&s);
                    } else if in_table {
                        current_cell.push_str(&s);
                    } else {
                        out.push_str(&s);
                    }
                }
                Tag::Image { dest_url, .. } => {
                    stats.images += 1;
                    let img_str = format!(
                        "\n#block(width: 100%, breakable: false)[\n  #align(center)[\n    #image(\"{}\", width: 100%)\n  ]\n]\n\n",
                        dest_url
                    );
                    if in_blockquote {
                        blockquote_content.push_str(&img_str);
                    } else if in_table {
                        current_cell.push_str(&format!("#image(\"{}\", width: 100%)", dest_url));
                    } else {
                        out.push_str(&img_str);
                    }
                }
                _ => {}
            },
            Event::End(tag_end) => match tag_end {
                TagEnd::Heading(_) => {
                    out.push_str("\n\n");
                }
                TagEnd::Paragraph => {
                    if in_blockquote {
                        blockquote_content.push_str("\n\n");
                    } else if !in_table {
                        out.push_str("\n\n");
                    }
                }
                TagEnd::BlockQuote(_) => {
                    in_blockquote = false;
                    stats.callouts += 1;
                    let trimmed = blockquote_content.trim();
                    let (border_color, bg_color, title, body) = if trimmed.starts_with("[!NOTE]") {
                        ("#0d6efd", "#f0f7ff", "ПРИМІТКА", trimmed.trim_start_matches("[!NOTE]").trim())
                    } else if trimmed.starts_with("[!WARNING]") {
                        ("#f59e0b", "#fffbeb", "УВАГА", trimmed.trim_start_matches("[!WARNING]").trim())
                    } else if trimmed.starts_with("[!TIP]") {
                        ("#10b981", "#f0fdf4", "ПОРАДА", trimmed.trim_start_matches("[!TIP]").trim())
                    } else if trimmed.starts_with("[!IMPORTANT]") {
                        ("#6366f1", "#eef2ff", "ВАЖЛИВО", trimmed.trim_start_matches("[!IMPORTANT]").trim())
                    } else if trimmed.starts_with("[!CAUTION]") {
                        ("#ef4444", "#fef2f2", "ЗАСТЕРЕЖЕННЯ", trimmed.trim_start_matches("[!CAUTION]").trim())
                    } else {
                        ("#94a3b8", "#f8fafc", "", trimmed)
                    };

                    out.push_str(&format!(
                        "\n#block(\n  width: 100%,\n  stroke: (left: 3pt + rgb(\"{}\"), rest: 0.5pt + rgb(\"#e2e8f0\")),\n  fill: rgb(\"{}\"),\n  radius: (right: 4pt),\n  inset: (x: 12pt, y: 10pt),\n  breakable: false\n)[\n",
                        border_color, bg_color
                    ));
                    if !title.is_empty() {
                        out.push_str(&format!(
                            "  #text(weight: \"bold\", fill: rgb(\"{}\"))[{}]\n  #v(0.2em)\n",
                            border_color, title
                        ));
                    }
                    out.push_str(&format!("  {}\n]\n\n", body));
                }
                TagEnd::CodeBlock => {
                    in_code_block = false;
                    out.push_str(&format!(
                        "\n#block(fill: rgb(\"#f8fafc\"), stroke: 0.5pt + rgb(\"#e2e8f0\"), inset: 8pt, radius: 3pt, width: 100%)[\n```{}\n{}\n```\n]\n\n",
                        code_lang, code_block_content.trim_end()
                    ));
                }
                TagEnd::List(_) => {
                    list_stack.pop();
                    out.push('\n');
                }
                TagEnd::Item => {
                    out.push('\n');
                }
                TagEnd::TableCell => {
                    current_row.push(current_cell.clone());
                }
                TagEnd::TableRow => {
                    table_cells.push(current_row.clone());
                }
                TagEnd::TableHead => {
                    _in_table_head = false;
                    table_cells.push(current_row.clone());
                }
                TagEnd::Table => {
                    in_table = false;
                    if !table_cells.is_empty() {
                        let num_cols = table_cells[0].len();
                        let cols_str = vec!["auto"; num_cols].join(", ");

                        out.push_str(&format!(
                            "\n#align(center)[\n#table(\n  columns: ({}),\n  fill: (col, row) => if row == 0 {{ rgb(\"#f1f5f9\") }} else if calc.even(row) {{ rgb(\"#f8fafc\") }} else {{ none }},\n  stroke: 0.5pt + rgb(\"#cbd5e1\"),\n  align: (col, row) => if row == 0 {{ center + horizon }} else {{ left + horizon }},\n  inset: (x: 8pt, y: 6pt),\n",
                            cols_str
                        ));

                        for (r_idx, row) in table_cells.iter().enumerate() {
                            for cell in row {
                                let c_trimmed = cell.trim();
                                if r_idx == 0 {
                                    out.push_str(&format!("  [*{}*],\n", c_trimmed));
                                } else {
                                    out.push_str(&format!("  [{}],\n", c_trimmed));
                                }
                            }
                        }
                        out.push_str(")\n]\n\n");
                    }
                }
                TagEnd::Emphasis => {
                    if in_blockquote {
                        blockquote_content.push('_');
                    } else if in_table {
                        current_cell.push('_');
                    } else {
                        out.push('_');
                    }
                }
                TagEnd::Strong => {
                    if in_blockquote {
                        blockquote_content.push('*');
                    } else if in_table {
                        current_cell.push('*');
                    } else {
                        out.push('*');
                    }
                }
                TagEnd::Strikethrough | TagEnd::Link => {
                    if in_blockquote {
                        blockquote_content.push(']');
                    } else if in_table {
                        current_cell.push(']');
                    } else {
                        out.push(']');
                    }
                }
                _ => {}
            },
            Event::Text(t) => {
                if in_code_block {
                    code_block_content.push_str(&t);
                } else {
                    let escaped = escape_typst_text(&t);
                    if in_blockquote {
                        blockquote_content.push_str(&escaped);
                    } else if in_table {
                        current_cell.push_str(&escaped);
                    } else {
                        out.push_str(&escaped);
                    }
                }
            }
            Event::Code(c) => {
                let code_inline = format!("`{}`", c);
                if in_blockquote {
                    blockquote_content.push_str(&code_inline);
                } else if in_table {
                    current_cell.push_str(&code_inline);
                } else {
                    out.push_str(&code_inline);
                }
            }
            Event::SoftBreak => {
                if in_blockquote {
                    blockquote_content.push(' ');
                } else if in_table {
                    current_cell.push(' ');
                } else {
                    out.push(' ');
                }
            }
            Event::HardBreak => {
                if in_blockquote {
                    blockquote_content.push_str(" \\\n");
                } else if in_table {
                    current_cell.push_str(" \\\n");
                } else {
                    out.push_str(" \\\n");
                }
            }
            Event::Rule => {
                out.push_str("\n#line(length: 100%, stroke: 0.5pt + rgb(\"#cbd5e1\"))\n\n");
            }
            _ => {}
        }
    }

    (out, stats)
}

/// Компілює Markdown безпосередньо у векторний PDF за допомогою рушія Typst
pub fn convert_markdown_to_pdf_typst(
    md_text: &str,
    output_path: &Path,
    source_dir: Option<&Path>,
    verbose: bool,
) -> Result<TypstStats, String> {
    let (typst_content, stats) = markdown_to_typst(md_text);

    if verbose {
        println!("  [v] Сгенеровано Typst-розмітку ({} байт)", typst_content.len());
        println!("  [v] Виявлено елементів: графіків={}, таблиць={}, callouts={}", 
            stats.images, stats.tables, stats.callouts);
    }

    let world = TypstWorld::new(typst_content, source_dir.map(Path::to_path_buf))?;

    if verbose {
        println!("  [v] Завантажено шрифтів у Typst World: {}", world.fonts.len());
    }

    let warned = typst::compile::<PagedDocument>(&world);
    for w in warned.warnings {
        if verbose {
            eprintln!("  [v] Typst warning: {:?}", w);
        }
    }

    let document = warned.output.map_err(|errors| {
        let err_msgs: Vec<String> = errors
            .into_iter()
            .map(|e| format!("{:?}: {}", e.span, e.message))
            .collect();
        format!("Помилка компіляції Typst:\n{}", err_msgs.join("\n"))
    })?;

    if verbose {
        println!("  [v] Typst успішно зверстав {} стор.", document.pages().len());
    }

    let pdf_bytes = pdf(&document, &PdfOptions::default())
        .map_err(|errors| format!("Помилка експорту PDF Typst: {:?}", errors))?;

    fs::write(output_path, &pdf_bytes)
        .map_err(|e| format!("Не вдалося зберегти PDF: {}", e))?;

    if verbose {
        println!("  [v] Збережено PDF: {} ({} байт)", output_path.display(), pdf_bytes.len());
    }

    Ok(stats)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_typst_compilation_minimal() {
        let md = "# Тестовий заголовок\n\nЦе перевірка генерації звітів українською мовою: [✓] ПІДТВЕРДЖЕНО.\n";
        let (typst_markup, _) = markdown_to_typst(md);
        let world = TypstWorld::new(typst_markup, None).expect("World init failed");
        let warned = typst::compile::<PagedDocument>(&world);
        assert!(warned.output.is_ok(), "Compilation failed: {:?}", warned.output.err());
        let doc = warned.output.unwrap();
        assert_eq!(doc.pages().len(), 1);
        let pdf_data = pdf(&doc, &PdfOptions::default()).expect("PDF export failed");
        assert!(!pdf_data.is_empty());
    }
}
