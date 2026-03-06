#!/bin/bash
set -e

BRANCH_NAME="$1"
COMMIT_MSG="$2"

if [ -z "$BRANCH_NAME" ] || [ -z "$COMMIT_MSG" ]; then
    echo "Not pushed :(. Usage: ./gitAUTOPUSH.sh branch_name \"commit_message\""
    exit 1
fi

GH_PATH=$(which gh)
if [ -z "$GH_PATH" ]; then
    echo "GitHub CLI (gh) not found. Install: brew install gh"
    exit 1
fi

git stash

git checkout main
git pull origin main

if git rev-parse --verify "$BRANCH_NAME" >/dev/null 2>&1; then
    git branch -D "$BRANCH_NAME"
fi
git push origin --delete "$BRANCH_NAME" 2>/dev/null || true

git checkout -b "$BRANCH_NAME"

git stash pop

if git diff-index --quiet HEAD --; then
    echo "Nothing to commit. Cleaned up and returned to main."
    git checkout main
    exit 0
fi

git add .
git commit -m "$COMMIT_MSG"
git push origin "$BRANCH_NAME"

$GH_PATH pr create \
  --base main \
  --head "$BRANCH_NAME" \
  --title "$COMMIT_MSG" \
  --body "pull request from gh CLI"


$GH_PATH pr merge "$BRANCH_NAME" --merge --delete-branch

git checkout main
git pull origin main

git branch -D "$BRANCH_NAME"

echo "Branch '$BRANCH_NAME' committed, PR created, merged, and cleaned up successfully."
