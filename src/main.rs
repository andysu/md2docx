mod cli;
mod converter;
mod mermaid;
mod template;

use clap::Parser;
use cli::Cli;
use std::path::PathBuf;
use std::process;
use std::time::Instant;

fn main() {
    let args = Cli::parse();

    let start = Instant::now();

    if !args.check && args.output.is_none() {
        eprintln!("Помилка: параметр -o/--output є обов'язковим, якщо не вказано --check");
        process::exit(1);
    }

    let rendered_text = match template::render_template(
        &args.template,
        &args.data,
        &args.var,
        args.keep_md.as_deref(),
    ) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("Помилка шаблонізації: {}", e);
            process::exit(1);
        }
    };

    if args.check {
        println!("=== Перевірка документа (Rust md2docx): {} ===", args.template.display());
        let dummy_out = PathBuf::from("/tmp/md2docx_dummy_check.docx");
        match converter::convert_markdown_to_docx(&rendered_text, &dummy_out, args.template.parent()) {
            Ok(stats) => {
                let _ = std::fs::remove_file(&dummy_out);
                println!("  ✓ Синтаксис валідний");
                println!("  Заголовків:        {}", stats.headings);
                println!("  Абзаців:           {}", stats.paragraphs);
                println!("  Таблиць:           {}", stats.tables);
                println!("  Блоків коду:       {}", stats.code_blocks);
                println!("  Діаграм Mermaid:   {}", stats.mermaid_diagrams);
                println!("  Callout-блоків:    {}", stats.callouts);
                println!("\nЧас перевірки: {:?}", start.elapsed());
                return;
            }
            Err(e) => {
                eprintln!("Помилка валідації: {}", e);
                process::exit(1);
            }
        }
    }

    let output_path = args.output.unwrap();
    match converter::convert_markdown_to_docx(&rendered_text, &output_path, args.template.parent()) {
        Ok(stats) => {
            println!("Готово: {} (за {:?})", output_path.display(), start.elapsed());
            if let Some(keep) = args.keep_md {
                println!("Заповнений MD збережено: {}", keep.display());
            }
            if stats.mermaid_diagrams > 0 {
                println!("Вбудовано діаграм Mermaid: {}", stats.mermaid_diagrams);
            }
            if stats.callouts > 0 {
                println!("Стилізовано Callout-блоків: {}", stats.callouts);
            }
        }
        Err(e) => {
            eprintln!("Помилка конвертації: {}", e);
            process::exit(1);
        }
    }
}
