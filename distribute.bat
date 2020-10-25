@echo off

echo Processing NMRA Release %1...
.\bin\Win64\NMRAMembers -s %1 > %1.log 2>&1
echo Done. Check %1.log for any warnings or errors

