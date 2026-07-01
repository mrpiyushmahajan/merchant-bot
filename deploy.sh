#!/bin/bash
# ─────────────────────────────────────────────────────────────
#  deploy.sh  —  Push merchant-bot to GitHub + open Railway
#  Usage: bash deploy.sh
# ─────────────────────────────────────────────────────────────
set -e

GITHUB_USER="mrpiyushmahajan"
REPO_NAME="merchant-bot"

echo ""
echo "═══════════════════════════════════════════════"
echo "  Merchant Bot  →  GitHub + Railway Deploy"
echo "═══════════════════════════════════════════════"
echo ""

# ── Step 1: Unpack the bundle ─────────────────────────────────
BUNDLE="$HOME/Downloads/merchant-bot.bundle"
DEST="$HOME/Projects/$REPO_NAME"

if [ ! -f "$BUNDLE" ]; then
  echo "❌  Bundle not found at $BUNDLE"
  echo "    Please move the merchant-bot.bundle file to ~/Downloads/ and re-run."
  exit 1
fi

if [ -d "$DEST" ]; then
  echo "📁  Folder $DEST already exists — skipping clone."
else
  echo "📦  Cloning from bundle → $DEST"
  git clone "$BUNDLE" "$DEST"
fi

cd "$DEST"

# ── Step 2: Create GitHub repo ────────────────────────────────
echo ""
echo "🐙  Creating GitHub repo: $GITHUB_USER/$REPO_NAME"

if command -v gh &>/dev/null; then
  # gh CLI available
  gh repo create "$REPO_NAME" \
    --public \
    --description "GPay + Paytm merchant CSV → Excel turnover report Telegram bot" \
    --source=. \
    --remote=origin \
    --push 2>/dev/null || true

  # If repo already exists just set remote and push
  git remote set-url origin "https://github.com/$GITHUB_USER/$REPO_NAME.git" 2>/dev/null || \
  git remote add origin "https://github.com/$GITHUB_USER/$REPO_NAME.git"
  git push -u origin main
else
  echo "  gh CLI not found — using git directly."
  echo "  (If push fails, first create the repo at https://github.com/new)"
  git remote set-url origin "https://github.com/$GITHUB_USER/$REPO_NAME.git" 2>/dev/null || \
  git remote add origin "https://github.com/$GITHUB_USER/$REPO_NAME.git"
  git push -u origin main
fi

REPO_URL="https://github.com/$GITHUB_USER/$REPO_NAME"
echo ""
echo "✅  Code is on GitHub: $REPO_URL"

# ── Step 3: Open Railway ──────────────────────────────────────
echo ""
echo "🚂  Opening Railway in your browser…"
echo "    1. Click  'New Project'"
echo "    2. Choose 'Deploy from GitHub repo'"
echo "    3. Select  $REPO_NAME"
echo "    4. Go to Variables tab → add:  BOT_TOKEN = <your token>"
echo "    5. Railway auto-deploys — done!"
echo ""

RAILWAY_URL="https://railway.app/new/github/$GITHUB_USER/$REPO_NAME"
open "$RAILWAY_URL" 2>/dev/null || xdg-open "$RAILWAY_URL" 2>/dev/null || \
  echo "    Open this URL manually: https://railway.app"

echo "═══════════════════════════════════════════════"
echo "  All done! Bot will be live in ~2 minutes"
echo "  after you add BOT_TOKEN in Railway."
echo "═══════════════════════════════════════════════"
echo ""
