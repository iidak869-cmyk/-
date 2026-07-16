@echo off
rem 契約数更新(NG含む)自動化の実行用バッチ。
rem タスクスケジューラからはこのバッチを起動する。
cd /d "%~dp0"
py -3 main.py
exit /b %ERRORLEVEL%
