#!/bin/zsh

dist_path="../bin/MacOSX";
work_path="./MacOSX";

if [ ! -d ${dist_path} ]; then
    mkdir -p ${dist_path};
fi;

if [ ! -d ${wokr_path} ]; then
    mkdir -p ${work_path};
fi;

for script in ../src/*.py; do
    echo "Building ${script}";
    echo "";
#    pyinstaller -y -F --distpath ${dist_path} --workpath ${work_path} ${script}
    pyinstaller -y -F --distpath ${dist_path} --workpath ${work_path} ${script} \
    --hidden-import pyexcel_xlsx \
    --hidden-import pyexcel_xlsx.xlsxr \
    --hidden-import pyexcel_xlsx.xlsxw \
    --hidden-import pyexcel_io.readers.csvr \
    --hidden-import pyexcel_io.readers.csvz \
    --hidden-import pyexcel_io.readers.tsv \
    --hidden-import pyexcel_io.readers.tsvz \
    --hidden-import pyexcel_io.writers.csvw \
    --hidden-import pyexcel_io.writers.csvz \
    --hidden-import pyexcel_io.writers.tsv \
    --hidden-import pyexcel_io.writers.tsvz \
    --hidden-import pyexcel_io.database.importers.django \
    --hidden-import pyexcel_io.database.importers.sqlalchemy \
    --hidden-import pyexcel_io.database.exporters.django
done;
echo "Compiled Programs:";
ls -l ${dist_path};

