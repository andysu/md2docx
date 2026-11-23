use regex::Regex;
use serde_json::{Map, Value};
use std::fs;
use std::path::{Path, PathBuf};

pub fn split_front_matter(text: &str) -> (Value, String) {
    let re = Regex::new(r"(?s)\A---\r?\n(.*?)\r?\n---\r?\n").unwrap();
    if let Some(caps) = re.captures(text) {
        let raw_yaml = caps.get(1).unwrap().as_str();
        let end_idx = caps.get(0).unwrap().end();
        let body = &text[end_idx..];
        let val: Value = serde_yaml::from_str(raw_yaml).unwrap_or(Value::Object(Map::new()));
        (val, body.to_string())
    } else {
        (Value::Object(Map::new()), text.to_string())
    }
}

pub fn deep_merge(base: &mut Value, override_val: &Value) {
    match (base, override_val) {
        (Value::Object(base_map), Value::Object(over_map)) => {
            for (k, v) in over_map {
                if let Some(base_entry) = base_map.get_mut(k) {
                    deep_merge(base_entry, v);
                } else {
                    base_map.insert(k.clone(), v.clone());
                }
            }
        }
        (base_slot, over_val) => {
            *base_slot = over_val.clone();
        }
    }
}

pub fn set_nested_value(data: &mut Value, dotted_key: &str, value: &str) {
    if !data.is_object() {
        *data = Value::Object(Map::new());
    }
    let parts: Vec<&str> = dotted_key.split('.').collect();
    let mut current = data;

    for &part in &parts[..parts.len() - 1] {
        if !current.is_object() {
            *current = Value::Object(Map::new());
        }
        current = current
            .as_object_mut()
            .unwrap()
            .entry(part)
            .or_insert_with(|| Value::Object(Map::new()));
    }

    if let Some(last) = parts.last() {
        if let Some(map) = current.as_object_mut() {
            // Спробуємо розпарсити як число/bool/json, інакше як звичайний рядок
            let parsed_val = serde_json::from_str::<Value>(value)
                .unwrap_or_else(|_| Value::String(value.to_string()));
            map.insert(last.to_string(), parsed_val);
        }
    }
}

pub fn load_data_files(paths: &[PathBuf]) -> Result<Value, String> {
    let mut merged = Value::Object(Map::new());
    for p in paths {
        if !p.exists() {
            return Err(format!("Data-файл не знайдено: {}", p.display()));
        }
        let content = fs::read_to_string(p)
            .map_err(|e| format!("Не вдалося прочитати {}: {}", p.display(), e))?;
        let chunk: Value = serde_yaml::from_str(&content)
            .map_err(|e| format!("Помилка YAML у {}: {}", p.display(), e))?;
        deep_merge(&mut merged, &chunk);
    }
    Ok(merged)
}

pub fn render_template(
    template_path: &Path,
    data_paths: &[PathBuf],
    cli_vars: &[String],
    keep_md: Option<&Path>,
) -> Result<String, String> {
    let raw = fs::read_to_string(template_path)
        .map_err(|e| format!("Не вдалося відкрити {}: {}", template_path.display(), e))?;

    let (mut context, body) = split_front_matter(&raw);
    let ext_data = load_data_files(data_paths)?;
    deep_merge(&mut context, &ext_data);

    for pair in cli_vars {
        if let Some((k, v)) = pair.split_once('=') {
            set_nested_value(&mut context, k.trim(), v.trim());
        }
    }

    let mut env = minijinja::Environment::new();
    env.set_trim_blocks(true);
    env.set_lstrip_blocks(true);

    let rendered = env
        .render_str(&body, &context)
        .map_err(|e| format!("Помилка шаблонізації Jinja2: {}", e))?;

    if let Some(keep_path) = keep_md {
        if let Some(parent) = keep_path.parent() {
            let _ = fs::create_dir_all(parent);
        }
        fs::write(keep_path, &rendered)
            .map_err(|e| format!("Не вдалося зберегти {}: {}", keep_path.display(), e))?;
    }

    Ok(rendered)
}
