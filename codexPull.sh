#!/bin/bash
set -e

# Switch to main only if no branch argument is given
if [ $# -eq 0 ]; then
    echo "No branch provided — switching to main..."
    git checkout main
else
    echo "Branch provided: $1 — skipping checkout of main."
fi

echo "Deleting all local branches except 'main'..."
git remote prune origin
git branch | grep -v "main" | xargs -r git branch -D
echo "Local branches deleted."

# Clean up stale remote-tracking branches
echo "Pruning remote-tracking branches..."
git fetch --prune

echo "Cleanup complete ✅"
