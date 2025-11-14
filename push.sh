#!/bin/bash

# Check if commit message is provided
if [ -z "$1" ]; then
    echo "Error: Please provide a commit message"
    echo "Usage: ./push.sh \"your commit message\""
    exit 1
fi

git add .
git commit -m "$1"
git push -u origin main