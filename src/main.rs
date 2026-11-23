mod cli;
mod converter;
mod mermaid;
mod template;

use clap::{CommandFactory, Parser};
use cli::Cli;
use std::path::PathBuf;
use std::process;
use std::time::Instant;

fn main() {
    // Якщо запущено без жодного аргументу — показати повну довідку
    if std::env::args().len() <= 1 {
        let mut cmd = Cli::command();
        let _ = cmd.print_help();
        println!();
        return;
    }

    let args = Cli::parse();

    let template_path = match &args.template {
        Some(p) => p,
        None => {
            let mut cmd = Cli::command();
            let _ = cmd.print_help();
            println!();
            return;
        }
    };

    let start = Instant::now();

    if !args.check && args.output.is_none() {
        eprintln!("Помилка: параметр -o/--output є обов'язковим, якщо не вказано --check");
        eprintln!("Спробуйте 'md2docx --help' для отримання довідки.");
        process::exit(1);
    }

    if args.verbose {
        println!("=== Запуск md2docx [verbose mode] ===");
        println!("  [v] Вхідний файл: {}", template_path.display());
    }

    let rendered_text = match template::render_template(
        template_path,
        &args.data,
        &args.var,
        args.keep_md.as_deref(),
        args.verbose,
    ) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("Помилка шаблонізації: {}", e);
            process::exit(1);
        }
    };

    if args.check {
        println!("=== Перевірка документа (Rust md2docx): {} ===", template_path.display());
        let dummy_out = PathBuf::from("/tmp/md2docx_dummy_check.docx");
        match converter::convert_markdown_to_docx(&rendered_text, &dummy_out, template_path.parent(), args.verbose) {
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
    match converter::convert_markdown_to_docx(&rendered_text, &output_path, template_path.parent(), args.verbose) {
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
