@echo off

set src_dir="..\src"
set dist_path="..\bin\Win64"
set work_path="..\build\Win64"
set spec_path="..\build\spec"

cd %SRC_DIR%

if not exist %dist_path% (mkdir %dist_path%)
if not exist %work_path% (mkdir %work_path%)
if not exist %spec_path% (mkdir %spec_path%)

for %%S in (*.py) do (
    echo "Building %%S..."
    echo " "
    pyinstaller --noconfirm --onefile --specpath %spec_path% --distpath %dist_path% --workpath %work_path% --hidden-import pyexcel_io.writers %%S
)
REM    --hidden-import pyexcel_io.readers.csvr^
REM    --hidden-import pyexcel_io.readers.csvz^
REM    --hidden-import pyexcel_io.readers.tsv^
REM    --hidden-import pyexcel_io.readers.tsvz^
REM    --hidden-import pyexcel_io.writers.csvw^
REM    --hidden-import pyexcel_io.readers.csvz^
REM    --hidden-import pyexcel_io.readers.tsv^
REM    --hidden-import pyexcel_io.readers.tsvz^
REM    --hidden-import pyexcel_io.database.importers.django^
REM    --hidden-import pyexcel_io.database.importers.sqlalchemy^
REM    --hidden-import pyexcel_io.database.exporters.django^
REM    --hidden-import pyexcel_io.database.exporters.sqlalchemy^
REM    --hidden-import pyexcel_xls^
REM    --hidden-import pyexcel_xls.xlsr^
REM    --hidden-import pyexcel_xls.xlsw^
REM    --hidden-import pyexcel.plugins^
REM    --hidden-import pyexcel.plugins.parsers^
REM    --hidden-import pyexcel.plugins.renderers^
REM    --hidden-import pyexcel.plugins.sources^
REM    --hidden-import pyexcel.plugins.sources.file_input^
REM    --hidden-import pyexcel.plugins.parsers.excel^
REM    --hidden-import pyexcel_xls^
REM    --hidden-import pyexcel_xls.xls^
REM    --hidden-import pyexcel_xlsx^
REM    --hidden-import pyexcel_xlsx.xlsx^
REM    --hidden-import pyexcel_xls^
REM    --hidden-import pyexcel_xls.xls^
echo "Compiled Programs:"
dir %dist_path%
cd ..

