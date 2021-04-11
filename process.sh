#!/bin/zsh

machine_arch=`uname -m`
echo "Processing NMRA Release ${1}...";
./bin/MacOSX/$%{machine_arch}/NMRAMembers $1 >& $1.log 
echo "Done. Check ${1}.log for any warnings or errors";

