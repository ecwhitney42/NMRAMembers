#!/bin/zsh

echo "Processing and Email Distributionn NMRA Release ${1}...";
machine_arch=`uname -m`
./bin/MacOSX/${machine_arch}/NMRAMembers -s $1 >& $1.log 
echo "Done. Check ${1}.log for any warnings or errors";

