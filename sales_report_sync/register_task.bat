@echo off
rem Windowsタスクスケジューラへ「毎日9:30実行」のタスクを登録する。
rem 管理者権限のコマンドプロンプトで一度だけ実行する。
schtasks /Create /TN "契約数更新_営業報告同期" /TR "\"%~dp0run_sync.bat\"" /SC DAILY /ST 09:30 /F
if %ERRORLEVEL% EQU 0 (
    echo タスクを登録しました: 毎日 9:30 に実行されます。
) else (
    echo タスクの登録に失敗しました。管理者権限で実行しているか確認してください。
)
pause
