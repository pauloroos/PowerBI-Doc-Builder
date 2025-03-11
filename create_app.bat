@echo off
cd /d %~dp0

REM Gera o executável com ícone, arquivos core, imagens e config incluídos
python -m PyInstaller --noconsole --onefile PowerBIDocBuilderApp.py ^
--icon=assets\\icon.ico ^
--add-data "config.json;." ^
--add-data "assets\\icon.png;assets" ^
--add-data "assets\\seu-logo.png;assets" ^
--add-data "assets\\image.png;assets" ^
--add-data "assets\\icon.ico;assets" ^
--add-data "core\\helpers.py;core" ^
--add-data "core\\ai_description.py;core" ^
--add-data "core\\diagram_generator.py;core" ^
--add-data "core\\pbi_extractor.py;core"

pause
