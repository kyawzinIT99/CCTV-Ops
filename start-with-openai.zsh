#!/bin/zsh
set -e
SCRIPT_DIR="$(cd -- "$(dirname -- "$0")" && pwd)"
cd "$SCRIPT_DIR"
read -s "OPENAI_API_KEY?Enter a NEW OpenAI API key (input hidden): "
print
if [[ -z "$OPENAI_API_KEY" ]]; then
  print -u2 "No key entered. The app was not started."
  exit 1
fi
export OPENAI_API_KEY
export OPENAI_MODEL="${OPENAI_MODEL:-gpt-5.6-luna}"
exec python3 server.py
