TARGETDIR	:= ../bin/MacOSX
TARGER		:= $(TARGETDIR)/NMRAMembers
WORKDIR		:= ./MacOSX
SRCDIR		:= ../src
HIDDEN		:= --hidden-import pyexcel_xlsx \
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


COMPILE	:= pyinstaller -y -F --distpath $(TARGETDIR) --workpath $(WORKDIR) 

$(TARGET): $(SRCDIR)/*.py
	$(COMPILE) $(inputs) $(HIDDEN)

#$(WORKDIR):
#	mkdir -p $(WORKDIR)

#$(TARGETDIR):
#	mkdir -p $(TARGETDIR)




