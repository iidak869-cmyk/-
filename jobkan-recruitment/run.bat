@echo off
REM タスクスケジューラから実行するためのバッチファイル
REM このファイルがあるフォルダに移動してから node を実行する
cd /d "%~dp0"
node src\index.js >> work\run.log 2>&1
