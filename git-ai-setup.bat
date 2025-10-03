@echo off
REM AI Git Analysis Setup Script for Windows

echo 🤖 Setting up AI-Powered Git Analysis...

REM Make the hook executable (Windows equivalent)
echo Setting up Git hooks...

REM Copy the hook to the correct location
if not exist ".git\hooks" mkdir ".git\hooks"

REM Create Windows-compatible version of the hook
echo #!/bin/bash > ".git\hooks\pre-commit"
type ".git\hooks\pre-commit" >> ".git\hooks\pre-commit.tmp"
move ".git\hooks\pre-commit.tmp" ".git\hooks\pre-commit"

REM Set executable permissions (if using Git Bash)
git update-index --chmod=+x ".git\hooks\pre-commit"

echo.
echo ✅ Git hooks installed successfully!
echo.
echo 📋 Setup Instructions:
echo 1. Set your OpenAI API key:
echo    set OPENAI_API_KEY=your_api_key_here
echo    OR add it to your environment variables
echo.
echo 2. Install required dependencies:
echo    npm install -g jq curl
echo.
echo 3. The hook will now run automatically on each commit
echo.
echo 🎉 AI Git Analysis is ready!

pause