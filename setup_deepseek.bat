@echo off
cd /d "%~dp0"
if not exist ".env" copy ".env.example" ".env" >nul
echo A private .env file is ready. Paste your DeepSeek API Key after LLM_API_KEY= and save.
echo Do not share this file or upload it to GitHub.
notepad ".env"
