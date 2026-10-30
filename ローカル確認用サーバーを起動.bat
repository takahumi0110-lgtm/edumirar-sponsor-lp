@echo off
cd /d "%~dp0tools"
echo ローカル確認用サーバーを起動します。
echo ブラウザで http://localhost:8080 を開いてください。
echo 終了するにはこのウィンドウを閉じるか Ctrl+C を押してください。
node local-server.js
pause
