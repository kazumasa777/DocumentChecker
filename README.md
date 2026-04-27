Aspose.Diagram オフライン・プラグイン配置手順（Windows配布用）

1. ネットワーク接続できる開発PCで、配布アプリと同じPythonバージョン/bit数の環境を用意します。
2. 次のようにローカルフォルダへ取得します。
   python -m pip download aspose-diagram-python -d aspose_download
3. 取得した wheel を展開し、aspose フォルダ一式を次の場所へ配置します。
   <exeと同じフォルダ>\plugins\aspose_diagram\aspose
4. 配置例:
   documentChecker_gui_new.exe
   plugins\aspose_diagram\aspose\diagram\...

補足:
- 実行PCでは pip install を行いません。
- 環境変数 ASPOSE_DIAGRAM_PLUGIN_DIR に配置先フォルダを指定することもできます。
- Python 3.14 の場合、Aspose側の対応wheelが存在しないと読み込めません。配布時はサポート済みPythonでビルドしてください。


documentChecker GUI Windows EXE化手順

1. 開発PCにPythonを用意します。
   オフライン実行PCではPython不要です。

2. このフォルダで build_exe.bat を実行します。

3. EXE作成後、以下のフォルダ一式を実行PCへコピーします。
   dist\documentChecker_gui_new\

4. Aspose.Diagramを同梱する場合
   開発PCのPythonに aspose-diagram-python を導入してから build_exe.bat を実行してください。
   batが aspose.diagram を検出すると --collect-all aspose でEXEフォルダへ同梱します。

5. Aspose.Diagramをプラグイン方式で後から配置する場合
   以下の場所へAsposeの aspose フォルダ一式を置いてください。

   dist\documentChecker_gui_new\plugins\aspose_diagram\aspose\diagram\...

6. Windows前提です。
   Linux/Mac用のsoffice/libreoffice探索は入れていません。




@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ============================================================
REM documentChecker GUI EXE build script for Windows
REM - Builds one-folder EXE with PyInstaller
REM - Aspose.Diagram is optional:
REM     1) If installed in the build Python, it is bundled by --collect-all aspose
REM     2) If not installed, build still continues and plugins\aspose_diagram can be copied later
REM ============================================================

cd /d "%~dp0"

set APP_NAME=documentChecker_gui_new
set ENTRY=documentChecker_gui_new.py
set PLUGIN_DIR=plugins
set ASPOSE_PLUGIN_DIR=plugins\aspose_diagram

if not exist "%ENTRY%" (
    echo [ERROR] %ENTRY% が見つかりません。このbatを documentChecker_gui_new.py と同じフォルダに置いて実行してください。
    pause
    exit /b 1
)

if not exist "documentChecker.py" (
    echo [ERROR] documentChecker.py が見つかりません。このbatを documentChecker.py と同じフォルダに置いて実行してください。
    pause
    exit /b 1
)

if not exist "%PLUGIN_DIR%" mkdir "%PLUGIN_DIR%"
if not exist "%ASPOSE_PLUGIN_DIR%" mkdir "%ASPOSE_PLUGIN_DIR%"

REM Python launcher check
py -3 --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python が見つかりません。開発PCで Python 3.x をインストールしてから実行してください。
    pause
    exit /b 1
)

REM Install PyInstaller only on build PC. Offline PCではこのbatを実行しない想定です。
py -3 -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [INFO] PyInstaller が見つからないため、開発PCにインストールします。
    py -3 -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] PyInstaller のインストールに失敗しました。
        pause
        exit /b 1
    )
)

REM Optional Aspose bundle detection
set ASPOSE_ARGS=
py -3 -c "import aspose.diagram" >nul 2>&1
if errorlevel 1 (
    echo [INFO] aspose.diagram はこのPython環境に未導入です。
    echo [INFO] EXE作成は継続します。Asposeを使う場合は配布フォルダに plugins\aspose_diagram を配置してください。
) else (
    echo [INFO] aspose.diagram を検出しました。EXEに同梱します。
    set ASPOSE_ARGS=--collect-all aspose
)

REM Clean previous build artifacts
if exist "build" rmdir /s /q "build"
if exist "dist\%APP_NAME%" rmdir /s /q "dist\%APP_NAME%"
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"

REM Build one-folder EXE. onefileではなくonedir推奨: 画像/Office/Aspose周辺の依存確認が容易なため。
py -3 -m PyInstaller ^
  --noconsole ^
  --onedir ^
  --clean ^
  --name "%APP_NAME%" ^
  --hidden-import documentChecker ^
  --hidden-import openpyxl ^
  --hidden-import docx ^
  --hidden-import pypdf ^
  --hidden-import fitz ^
  --hidden-import PIL ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --add-data "plugins;plugins" ^
  !ASPOSE_ARGS! ^
  "%ENTRY%"

if errorlevel 1 (
    echo [ERROR] EXE化に失敗しました。上記ログを確認してください。
    pause
    exit /b 1
)

REM Ensure plugin folder exists in dist even when empty
if not exist "dist\%APP_NAME%\plugins" mkdir "dist\%APP_NAME%\plugins"
if not exist "dist\%APP_NAME%\plugins\aspose_diagram" mkdir "dist\%APP_NAME%\plugins\aspose_diagram"

REM Copy README files when present
if exist "ASPOSE_PLUGIN_README.txt" copy /Y "ASPOSE_PLUGIN_README.txt" "dist\%APP_NAME%\ASPOSE_PLUGIN_README.txt" >nul
if exist "RUN_README.txt" copy /Y "RUN_README.txt" "dist\%APP_NAME%\RUN_README.txt" >nul

echo.
echo [OK] EXE作成が完了しました。
echo     dist\%APP_NAME%\%APP_NAME%.exe
echo.
echo [NOTE] オフライン配布時は dist\%APP_NAME% フォルダ一式をコピーしてください。
echo [NOTE] Asposeを外部プラグインで使う場合は、以下に aspose フォルダ一式を配置してください。
echo     dist\%APP_NAME%\plugins\aspose_diagram\aspose\diagram\...
echo.
pause
exit /b 0
