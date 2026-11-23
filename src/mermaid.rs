use base64::Engine;
use regex::Regex;
use sha2::{Digest, Sha256};
use std::fs;
use std::path::PathBuf;
use std::process::Command;

pub fn sanitize_mermaid(code: &str) -> String {
    let mut clean = code.trim().to_string();
    let re_comment = Regex::new(r"(?m)^([ \t]*(?:graph|flowchart)\s+[A-Za-z]+)%%").unwrap();
    clean = re_comment.replace_all(&clean, "$1\n%%").to_string();

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

pub fn render_mermaid(code: &str, verbose: bool) -> Option<Vec<u8>> {
    let clean = sanitize_mermaid(code);
    let mut hasher = Sha256::new();
    hasher.update(clean.as_bytes());
    let hash_hex = hex::encode(&hasher.finalize()[..10]);

    let cache_dir = get_cache_dir();
    let _ = fs::create_dir_all(&cache_dir);
    let cache_file = cache_dir.join(format!("mermaid_{}.png", hash_hex));

    if let Ok(bytes) = fs::read(&cache_file) {
        if bytes.len() > 100 {
            if verbose {
                println!("  [v] Діаграма Mermaid [{}] знайдена в локальному кеші ({} байт)", hash_hex, bytes.len());
            }
            return Some(bytes);
        }
    }

    if verbose {
        println!("  [v] Діаграма Mermaid [{}] відсутня в кеші, запит до Kroki API...", hash_hex);
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
            if verbose {
                println!("  [v] Kroki API успішно згенерував PNG [{}] ({} байт)", hash_hex, bytes.len());
            }
            return Some(bytes);
        }
    }

    // Спроба 2: mermaid.ink fallback
    if verbose {
        println!("  [v] Запасний виклик mermaid.ink для [{}]...", hash_hex);
    }
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
            if verbose {
                println!("  [v] mermaid.ink успішно згенерував зображення [{}]", hash_hex);
            }
            return Some(bytes);
        }
    }

    // Спроба 3: локальний mmdc
    let tmp_file = cache_dir.join(format!("temp_{}.mmd", hash_hex));
    if fs::write(&tmp_file, &clean).is_ok() {
        if verbose {
            println!("  [v] Спроба локального виклику mmdc для [{}]...", hash_hex);
        }
        let status = Command::new("mmdc")
            .args(["-i", tmp_file.to_str().unwrap(), "-o", cache_file.to_str().unwrap(), "-b", "white"])
            .output();
        let _ = fs::remove_file(tmp_file);
        if let Ok(out) = status {
            if out.status.success() {
                if let Ok(bytes) = fs::read(&cache_file) {
                    if bytes.len() > 100 {
                        if verbose {
                            println!("  [v] mmdc успішно відрендерив PNG [{}]", hash_hex);
                        }
                        return Some(bytes);
                    }
                }
            }
        }
    }

    if verbose {
        eprintln!("  [v] Попередження: не вдалося відрендерити діаграму [{}]", hash_hex);
    }
    None
}
