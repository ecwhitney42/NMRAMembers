#!/bin/zsh

src_dir="../src";
dist_path="../bin/MacOSX";
work_path="../build/MacOSX";
spec_path="../build/spec";

cd ${src_dir}

if [ ! -d ${dist_path} ]; then
    mkdir -p ${dist_path};
fi;

if [ ! -d ${work_path} ]; then
    mkdir -p ${work_path};
fi;

if [ ! -d ${spec_path} ]; then
    mkdir -p ${spec_path};
fi;

for script in *.py; do
    echo "Building ${script}";
    echo "";
    pyinstaller --noconfirm --onefile --specpath ${spec_path} --distpath ${dist_path} --workpath ${work_path} --hidden-import pyexcel_xls.xls ${script};
#    --hidden-import pyexcel_xlsx \
#    --hidden-import pyexcel_xlsx.xlsxr \
#    --hidden-import pyexcel_xlsx.xlsxw \
#    --hidden-import pyexcel_io.readers.csvr \
#    --hidden-import pyexcel_io.readers.csvz \
#    --hidden-import pyexcel_io.readers.tsv \
#    --hidden-import pyexcel_io.readers.tsvz \
#    --hidden-import pyexcel_io.writers.csvw \
#    --hidden-import pyexcel_io.writers.csvz \
#    --hidden-import pyexcel_io.writers.tsv \
#    --hidden-import pyexcel_io.writers.tsvz \
#    --hidden-import pyexcel_io.database.importers.django \
#    --hidden-import pyexcel_io.database.importers.sqlalchemy \
#    --hidden-import pyexcel_io.database.exporters.django \
done;
echo "Compiled Programs:";
ls -l ${dist_path};

