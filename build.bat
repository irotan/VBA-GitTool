@echo off
pushd "%~dp0"

echo ========================================================
echo  VbaGitTool.exe をビルド（コンパイル）します...
echo ========================================================
echo.

set "CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" (
    set "CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

if not exist "%CSC_PATH%" (
    echo [エラー] Windows標準の C# コンパイラ（csc.exe）が見つかりませんでした。
    popd
    pause
    exit /b 1
)

:: 読み込むファイルを VbaGitTool.cs に変更
"%CSC_PATH%" /target:winexe /out:VbaGitTool.exe /utf8output VbaGitTool.cs
if %errorlevel% neq 0 (
    echo.
    echo [エラー] コンパイルに失敗しました。上記のエラーメッセージを確認してください。
    popd
    pause
    exit /b 1
)

echo.
echo ========================================================
echo  成功！ VbaGitTool.exe が生成されました！
echo ========================================================
echo.
popd
pause