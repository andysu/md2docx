use clap::Parser;
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(
    name = "md2docx",
    about = "Універсальний високошвидкісний конвертер Markdown у DOCX (Mermaid, Callouts, OMML)",
    version
)]
pub struct Cli {
    /// Вхідний .md файл (шаблон або звичайний markdown)
    pub template: PathBuf,

    /// Шлях до вихідного .docx файлу
    #[arg(short, long)]
    pub output: Option<PathBuf>,

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
}
