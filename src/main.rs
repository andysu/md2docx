mod cli;
mod converter;
mod mermaid;
mod pdf_converter;
mod template;
mod typst_backend;

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
    let ext = output_path.extension().and_then(|s| s.to_str()).unwrap_or("").to_lowercase();

    if ext == "pdf" {
        if args.verbose {
            println!("  [v] Пряма генерація PDF через вбудований векторний рушій Typst");
        }
        match typst_backend::convert_markdown_to_pdf_typst(
            &rendered_text,
            &output_path,
            template_path.parent(),
            args.verbose,
        ) {
            Ok(stats) => {
                println!("Готово: {} (за {:?})", output_path.display(), start.elapsed());
                if stats.images > 0 {
                    println!("Вбудовано графіків/зображень: {}", stats.images);
                }
                if stats.tables > 0 {
                    println!("Сформовано таблиць: {}", stats.tables);
                }
                if stats.callouts > 0 {
                    println!("Стилізовано блоків висновків (Callouts): {}", stats.callouts);
                }
            }
            Err(e) => {
                eprintln!("Помилка генерації PDF через Typst: {}", e);
                process::exit(1);
            }
        }

        if let Some(keep) = args.keep_md {
            println!("Заповнений MD збережено: {}", keep.display());
        }
    } else {
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

                if args.pdf {
                    let mut pdf_out = output_path.clone();
                    pdf_out.set_extension("pdf");
                    if args.verbose {
                        println!("  [v] Генерація супутнього PDF через Typst: {}", pdf_out.display());
                    }
                    match typst_backend::convert_markdown_to_pdf_typst(
                        &rendered_text,
                        &pdf_out,
                        template_path.parent(),
                        args.verbose,
                    ) {
                        Ok(_) => {
                            println!("Також згенеровано PDF через Typst: {}", pdf_out.display());
                        }
                        Err(e) => {
                            eprintln!("Попередження: не вдалося згенерувати супутній PDF: {}", e);
                        }
                    }
                }
            }
            Err(e) => {
                eprintln!("Помилка конвертації: {}", e);
                process::exit(1);
            }
        }
    }
}

#[allow(dead_code)]
fn convert_docx_to_pdf_with_soffice(
    docx_path: &std::path::Path,
    pdf_output_path: &std::path::Path,
    verbose: bool,
) -> Result<(), String> {
    let has_soffice = std::process::Command::new("which")
        .arg("soffice")
        .output()
        .map(|o| o.status.success())
        .unwrap_or(false);

    if !has_soffice {
        return Err("Утиліту 'soffice' (LibreOffice) не знайдено в системі".to_string());
    }

    let mut unique_dir = std::env::temp_dir();
    unique_dir.push(format!("soffice_pdf_{}", std::process::id()));
    let _ = std::fs::create_dir_all(&unique_dir);

    if verbose {
        println!("  [v] Запуск LibreOffice (soffice) для бездоганної пагінації DOCX -> PDF...");
    }

    let status = std::process::Command::new("soffice")
        .args(&[
            "--headless",
            "--convert-to",
            "pdf",
            docx_path.to_str().unwrap(),
            "--outdir",
            unique_dir.to_str().unwrap(),
        ])
        .status()
        .map_err(|e| format!("Помилка виконання soffice: {}", e))?;

    if !status.success() {
        let _ = std::fs::remove_dir_all(&unique_dir);
        return Err(format!("soffice завершився з кодом помилки: {:?}", status.code()));
    }

    let file_stem = docx_path.file_stem().and_then(|s| s.to_str()).unwrap_or("document");
    let generated_pdf = unique_dir.join(format!("{}.pdf", file_stem));

    if generated_pdf.exists() {
        std::fs::copy(&generated_pdf, pdf_output_path)
            .map_err(|e| format!("Не вдалося скопіювати згенерований PDF: {}", e))?;
        let _ = std::fs::remove_dir_all(&unique_dir);
        Ok(())
    } else {
        let _ = std::fs::remove_dir_all(&unique_dir);
        Err("soffice завершився, але згенерований PDF файл не знайдено".to_string())
    }
}
