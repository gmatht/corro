#!/usr/bin/env bash
# Run the newest target/*/corro if it is newer than all inputs under src/ and
# ./Cargo.*; otherwise `cargo run --`. Forwards all arguments in either case.
set -euo pipefail
cd "$(dirname "$0")"

shopt -s nullglob

file_mtime() {
  case "$(uname -s 2>/dev/null || true)" in
  Darwin) stat -f %m "$1" ;;
  *)      stat -c %Y "$1" ;;
  esac
}

max_source_mtime=0
if [[ -d src ]]; then
  while IFS= read -r -d '' f; do
    m=$(file_mtime "$f")
    if ((m > max_source_mtime)); then max_source_mtime=$m; fi
  done < <(find src -type f -print0 2>/dev/null)
fi
for f in ./Cargo.*; do
  [[ -f $f ]] || continue
  m=$(file_mtime "$f")
  if ((m > max_source_mtime)); then max_source_mtime=$m; fi
done

newest_bin=
newest_m=0
for f in target/*/corro; do
  [[ -f $f ]] || continue
  m=$(file_mtime "$f")
  if [[ -z $newest_bin ]] || ((m > newest_m)); then
    newest_m=$m
    newest_bin=$f
  fi
done

if [[ -n $newest_bin ]] && ((newest_m > max_source_mtime)); then
  exec "$newest_bin" "$@"
else
  exec cargo run -- "$@"
fi
