#!/bin/zsh

echo "Processing NMRA Release ${1}...";
./bin/MacOS/NMRAMembers $1 >& $1.log 
echo "Done. Check ${1}.log for any warnings or errors";

