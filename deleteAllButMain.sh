#!/bin/bash
set -e

# Always start from main
echo "Switching to main branch..."
git checkout main

# Record current timestamped log file
LOG_FILE="deleted_branches_$(date +%Y%m%d_%H%M%S).log"

echo "Recording and deleting all local branches except 'main'..."
git branch | grep -v "main" | tee "$LOG_FILE" | xargs -r git branch -D

echo "Local branches deleted. Log saved to $LOG_FILE"

# Clean up stale remote-tracking branches
echo "Pruning remote-tracking branches..."
git fetch --prune

echo "Cleanup complete ✅"

