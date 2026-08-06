#!/usr/bin/env bash
# Erstellt das GitHub-Repo und pusht dieses Projekt (einmalig, mit deinem Login).
set -euo pipefail

REPO_NAME="${REPO_NAME:-WAMAS-Lift-Store-Angebot}"
REPO_DESC="${REPO_DESC:-WAMAS Lift & Store Angebot – SSI SCHÄFER Kalkulator & Angebotsgenerator}"
VISIBILITY="${VISIBILITY:-private}"   # private | public
OWNER="${OWNER:-}"                    # leer = dein GitHub-User; oder z.B. msobona

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if ! command -v gh >/dev/null 2>&1; then
  echo "Bitte GitHub CLI installieren: https://cli.github.com/"
  exit 1
fi

if ! gh auth status >/dev/null 2>&1; then
  echo "Bitte zuerst anmelden:  gh auth login"
  exit 1
fi

TARGET="$REPO_NAME"
if [[ -n "$OWNER" ]]; then
  TARGET="$OWNER/$REPO_NAME"
fi

echo "→ Erstelle Repository: $TARGET ($VISIBILITY)"
if [[ "$VISIBILITY" == "public" ]]; then
  VIS_FLAG=--public
else
  VIS_FLAG=--private
fi

# Repo anlegen (scheitert harmlos, falls es schon existiert)
gh repo create "$TARGET" $VIS_FLAG --description "$REPO_DESC" --source=. --remote=origin --push \
  || {
    echo "Repo existiert vermutlich schon – setze Remote und pushe…"
    if [[ -n "$OWNER" ]]; then
      URL="https://github.com/${OWNER}/${REPO_NAME}.git"
    else
      ME="$(gh api user --jq .login)"
      URL="https://github.com/${ME}/${REPO_NAME}.git"
    fi
    git remote remove origin 2>/dev/null || true
    git remote add origin "$URL"
    git push -u origin HEAD:master
  }

echo
echo "Fertig."
gh repo view "$TARGET" --web 2>/dev/null || true
echo "Repo-URL:"
gh repo view "$TARGET" --json url --jq .url
