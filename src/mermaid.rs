use base64::Engine;
use regex::Regex;
use sha2::{Digest, Sha256};
use std::fs;
use std::path::PathBuf;
use std::process::Command;

pub fn sanitize_mermaid(code: &str) -> String {
    let mut clean = code.trim().to_string();
    // 1. Розділити приліплені коментарі %% після напрямку (graph TD%% -> graph TD\n%%)
    let re_comment = Regex::new(r"(?m)^([ \t]*(?:graph|flowchart)\s+[A-Za-z]+)%%").unwrap();
    clean = re_comment.replace_all(&clean, "$1\n%%").to_string();

    // 2. Сумісність для застарілих gitGraph
    if clean.contains("gitGraph") {
        let re_options = Regex::new(r"(?s)options\s*\{[^}]*\}\s*end\s*").unwrap();
        clean = re_options.replace_all(&clean, "").to_string();
        if clean.contains("checkout master") && !clean.contains("branch master") {
            clean = clean.replace("checkout master", "checkout main");
        }
    }

    clean
}

pub fn get_cache_dir() -> PathBuf {
    if let Ok(home) = std::env::var("HOME") {
        PathBuf::from(home).join(".cache").join("md2docx").join("mermaid")
    } else {
        std::env::temp_dir().join("md2docx_mermaid_cache")
    }
}

pub fn render_mermaid(code: &str) -> Option<Vec<u8>> {
    let clean = sanitize_mermaid(code);
    let mut hasher = Sha256::new();
    hasher.update(clean.as_bytes());
    let hash_hex = hex::encode(&hasher.finalize()[..10]);

    let cache_dir = get_cache_dir();
    let _ = fs::create_dir_all(&cache_dir);
    let cache_file = cache_dir.join(format!("mermaid_{}.png", hash_hex));

    if let Ok(bytes) = fs::read(&cache_file) {
        if bytes.len() > 100 {
            return Some(bytes);
        }
    }

    // Спроба 1: Kroki POST API
    let kroki_res = ureq::post("https://kroki.io/mermaid/png")
        .set("Content-Type", "text/plain; charset=utf-8")
        .set("User-Agent", "md2docx-rs")
        .timeout(std::time::Duration::from_secs(12))
        .send_bytes(clean.as_bytes());

    if let Ok(resp) = kroki_res {
        let mut bytes = Vec::new();
        if resp.into_reader().read_to_end(&mut bytes).is_ok() && bytes.len() > 100 {
            let _ = fs::write(&cache_file, &bytes);
            return Some(bytes);
        }
    }

    // Спроба 2: mermaid.ink fallback
    let state = serde_json::json!({
        "code": clean,
        "mermaid": { "theme": "default" }
    });
    let b64 = base64::engine::general_purpose::STANDARD.encode(state.to_string().as_bytes());
    let url = format!("https://mermaid.ink/img/{}", b64);

    let ink_res = ureq::get(&url)
        .set("User-Agent", "md2docx-rs")
        .timeout(std::time::Duration::from_secs(10))
        .call();

    if let Ok(resp) = ink_res {
        let mut bytes = Vec::new();
        if resp.into_reader().read_to_end(&mut bytes).is_ok() && bytes.len() > 100 {
            let _ = fs::write(&cache_file, &bytes);
            return Some(bytes);
        }
    }

    // Спроба 3: локальний mmdc якщо встановлений
    let tmp_file = cache_dir.join(format!("temp_{}.mmd", hash_hex));
    if fs::write(&tmp_file, &clean).is_ok() {
        let status = Command::new("mmdc")
            .args(["-i", tmp_file.to_str().unwrap(), "-o", cache_file.to_str().unwrap(), "-b", "white"])
            .output();
        let _ = fs::remove_file(tmp_file);
        if let Ok(out) = status {
            if out.status.success() {
                if let Ok(bytes) = fs::read(&cache_file) {
                    if bytes.len() > 100 {
                        return Some(bytes);
                    }
                }
            }
        }
    }

    None
}
