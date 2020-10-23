# NMRAMembers

NMRAMembers is a Python 3 program used to take the NMRA Membership ZIP file that is sent out each month by national to the region and distribute all of the individual division member rosters to a representative from each division.

The Python program can be run from the command line using python3 or it can be compiled using pyinstaller for a user's platform.

At the current time, this has only been tested on MacOS X, however, it should be possible to run under Windows if the correct Python environment is set up.

NMRAMembers Python Installation Notes

Install Python 3.7.6 using the installer

pip3 install —upgrade pip

pip3 install xlrd

pip3 install xlwt

pip3 install pyexcel

pip3 install pyexcel_xlsx

pip3 install pyinstaller


For the Mac, there is an Apple Automator app that can be used do drag and drop the .zip file onto which executes the script.

