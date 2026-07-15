@echo off
cd /d "%~dp0"
if not exist logs mkdir logs
set PORT=3001
node server.js >> logs\server.out.log 2>> logs\server.err.log
