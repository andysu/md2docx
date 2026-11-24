#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== [1/3] Компіляція релізного бінарника md2docx (Cargo) ==="
cargo build --release

BIN_PATH="target/release/md2docx"
if [ ! -f "$BIN_PATH" ]; then
    echo "Помилка: бінарник $BIN_PATH не знайдено!"
    exit 1
fi

SIZE_BEFORE=$(stat -c%s "$BIN_PATH" 2>/dev/null || stat -f%z "$BIN_PATH")
SIZE_BEFORE_MB=$(awk -v s="$SIZE_BEFORE" 'BEGIN { printf "%.2f", s / (1024 * 1024) }')
echo "  [*] Розмір до компресії: ${SIZE_BEFORE_MB} МБ (${SIZE_BEFORE} байт)"

echo "=== [2/3] Застосування максимальної UPX компресії (--best --lzma) ==="
if command -v upx >/dev/null 2>&1; then
    upx --best --lzma --force-overwrite "$BIN_PATH"
else
    echo "Попередження: утиліту upx не знайдено на системі. Пропуск компресії."
fi

SIZE_AFTER=$(stat -c%s "$BIN_PATH" 2>/dev/null || stat -f%z "$BIN_PATH")
SIZE_AFTER_MB=$(awk -v s="$SIZE_AFTER" 'BEGIN { printf "%.2f", s / (1024 * 1024) }')
RATIO=$(awk -v b="$SIZE_BEFORE" -v a="$SIZE_AFTER" 'BEGIN { printf "%.1f", (1.0 - a / b) * 100 }')

echo "=== [3/3] Результат збірки ==="
echo "  [*] Розмір після компресії: ${SIZE_AFTER_MB} МБ (${SIZE_AFTER} байт)"
echo "  [*] Ефективність стиснення: -${RATIO}%"
echo "  [*] Бінарник готовий: $BIN_PATH"
