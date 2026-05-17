#!/bin/bash
# Diagnostic wrapper for ms365 MCP server
# Logs startup context to help debug failures when launched from non-project directories

SERVER_DIR="/home/zow/projects/ms-365-mcp-server"
LOG="$SERVER_DIR/logs/mcp-startup.log"

{
  echo "=== MS365 MCP Server Startup ==="
  echo "Timestamp: $(date -Iseconds)"
  echo "PID: $$"
  echo "Initial CWD: $(pwd)"
  echo "SHELL: $SHELL"
  echo "PATH (first 3): $(echo "$PATH" | tr ':' '\n' | head -3 | tr '\n' ':')"
  echo ""
  echo "--- Key files check ---"
  for f in "$SERVER_DIR/package.json" \
           "$SERVER_DIR/src/index.ts" \
           "$SERVER_DIR/src/endpoints.json" \
           "$SERVER_DIR/.token-cache.json" \
           "$SERVER_DIR/.env" \
           "$SERVER_DIR/node_modules/.bin/tsx"; do
    if [ -f "$f" ]; then
      echo "  EXISTS: $f"
    else
      echo "  MISSING: $f"
    fi
  done
  echo ""
  echo "--- Node/tsx versions ---"
  node --version 2>&1 || echo "  node: NOT FOUND"
  "$SERVER_DIR/node_modules/.bin/tsx" --version 2>&1 || echo "  tsx: NOT FOUND"
  echo ""
  echo "--- Environment (MS365/MCP-relevant) ---"
  env | grep -iE '^(MS365|MCP|NODE|DOTENV|HOME|USER|DISPLAY)' | sort
  echo ""
  echo "Changing CWD to: $SERVER_DIR"
} >> "$LOG" 2>&1

cd "$SERVER_DIR" || {
  echo "FATAL: Cannot cd to $SERVER_DIR" >> "$LOG"
  exit 1
}

echo "CWD after cd: $(pwd)" >> "$LOG"
echo "--- Launching tsx ---" >> "$LOG"
echo "" >> "$LOG"

# exec replaces this shell with tsx, so stdio passes through cleanly
exec "$SERVER_DIR/node_modules/.bin/tsx" "$SERVER_DIR/src/index.ts" "$@" 2>> "$LOG"
