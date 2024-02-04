#!/bin/zsh

month=${1};

#echo "Processing and Email Distributionn NMRA Release ${1}...";
#machine_arch=`uname -m`
#./bin/MacOSX/${machine_arch}/NMRAMembers -s $1 >& $1.log 
#echo "Done. Check ${1}.log for any warnings or errors";

echo "Distributing NMRA Roster Files for ${month}...";
osascript src/distribute.scpt ${month};
echo "Done."


