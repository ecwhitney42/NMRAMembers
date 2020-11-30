@echo off

set src_dir="..\src"
set dist_path="..\bin\Win64"
set work_path="..\build\Win64"
set spec_path="..\build\spec"

cd %SRC_DIR%

mkdir %dist_path%
mkdir %work_path%
mkdir %spec_path%

for %%S in (*.py) do (
    echo "Building %%S..."
    echo " "
    pyinstaller --noconfirm --onefile --specpath %spec_path% --distpath %dist_path% --workpath %work_path% --hidden-import pyexcel_xls.xls %%S
)

echo "Compiled Programs:"
dir %dist_path%

