use clap::Parser;
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(
    name = "md2docx",
    about = "Універсальний високошвидкісний конвертер Markdown у DOCX (Mermaid, Callouts, OMML)",
    version,
    after_help = "ПРИКЛАДИ:\n  \
        # Звичайна конвертація\n  \
        md2docx document.md -o output.docx\n\n  \
        # Детальний verbose-вивід конвеєра\n  \
        md2docx document.md -o output.docx -v\n\n  \
        # З підстановкою даних із YAML-файлу\n  \
        md2docx template.md -o output.docx --data data.yaml\n\n  \
        # Перевизначення змінних через CLI\n  \
        md2docx template.md -o output.docx --var client.name=\"ТОВ Ромашка\"\n\n  \
        # Швидка перевірка документа без запису .docx\n  \
        md2docx document.md --check\n"
)]
pub struct Cli {
    /// Вхідний .md файл (шаблон або звичайний markdown)
    pub template: Option<PathBuf>,

    /// Шлях до вихідного файлу (.docx або .pdf)
    #[arg(short, long)]
    pub output: Option<PathBuf>,

    /// Додатково генерувати автономний PDF файл поруч із DOCX
    #[arg(long)]
    pub pdf: bool,

    /// YAML data-файли зі змінними (можна вказати кілька разів)
    #[arg(long, value_name = "FILE")]
    pub data: Vec<PathBuf>,

    /// Окреме значення змінної, крапкова нотація (напр. --var client.name="Foo")
    #[arg(long, value_name = "KEY=VAL")]
    pub var: Vec<String>,

    /// Зберегти заповнений проміжний .md файл перед конвертацією
    #[arg(long, value_name = "PATH")]
    pub keep_md: Option<PathBuf>,

    /// Тільки перевірити документ і надрукувати звіт без збереження .docx
    #[arg(long)]
    pub check: bool,

    /// Детальний вивід процесу виконання (verbose)
    #[arg(short, long)]
    pub verbose: bool,
}
