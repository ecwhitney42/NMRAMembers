#!/bin/zsh

echo "Processing and Email Distributionn NMRA Release ${1}...";
./bin/MacOS/NMRAMembers -s $1 >& $1.log 
echo "Done. Check ${1}.log for any warnings or errors";

