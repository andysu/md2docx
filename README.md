# md2docx

Універсальний автономний конвертер Markdown у формат DOCX (Microsoft Word) із підтримкою підстановки змінних (Jinja2 + YAML), векторних математичних формул (LaTeX → OMML), графічних діаграм Mermaid та інформаційних блоків у стилі macOS.

## Ключові можливості

* **Один автономний файл:** `md2docx.py` містить усе необхідне й не потребує встановлення важких конвертаційних фреймворків.
* **Діаграми Mermaid:** автоматичний рендеринг блоків ```` ```mermaid ```` у високоякісні зображення з масштабуванням під ширину сторінки (використовує Kroki POST API з локальним кешуванням).
* **Callout-блоки (macOS SF Style):** підтримка GitHub Alerts (`> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`, `> [!CAUTION]`) з відповідними кольоровими лінійками, фоном та системними піктограмами.
* **Нативні формули Word:** трансляція LaTeX-формул у нативні редаговані об'єкти Word (OMML через Pandoc), а не у растрові картинки.
* **Незалежна нумерація списків:** гарантований перезапуск нумерації кожного списку з 1 через нативні механізми OpenXML (`startOverride`).
* **Шаблонізація:** підтримка змінних Jinja2 всередині Markdown, завантаження даних із YAML-файлів та перевизначення значень із командного рядка (`--var`).
* **Аудит документа:** прапорець `--check` для швидкої перевірки змінних, картинок і структури без збереження docx.

## Встановлення

```bash
git clone https://github.com/andysu/md2docx.git
cd md2docx
pip install -r requirements.txt
```

Для конвертації LaTeX-формул потрібен системний `pandoc`:
```bash
# Ubuntu / Debian
sudo apt install pandoc

# Fedora / RHEL
sudo dnf install pandoc

# macOS
brew install pandoc
```

## Використання

```bash
# 1. Звичайна конвертація
python3 md2docx.py document.md -o output.docx

# 2. З підстановкою даних із YAML
python3 md2docx.py template.md -o output.docx --data data.yaml

# 3. Перевизначення змінних через CLI
python3 md2docx.py template.md -o output.docx --data data.yaml --var client.name="ТОВ Світанок"

# 4. Перевірка документа без генерації файлу
python3 md2docx.py document.md --check
```

## Ліцензія

MIT
