# NMRAMembers

NMRAMembers is a Python 3 program used to take the NMRA Membership ZIP file that is sent out each month by national to the region and distribute all of the individual division member rosters to a representative from each division.

The Python program can be run from the command line using python3 or it can be compiled using pyinstaller for a user's platform.

At the current time, this has only been tested on MacOS X, however, it should be possible to run under Windows if the correct Python environment is set up.

## Installation

NOTE: The installation process is still in development as the author learn more about how to use these tools. At present, it is assumed that the binary program will be distributed and not the source files but since Windows support hasn't been added, Windows users will have to work out the Python installation and run the Python source code until Pyinstaller can be worked out on that platform.

### MacOSX

**Dependencies (Install PYthon 3):**

Install Python 3.7.6 using the installer. The reason this specific version is beause the current version of pyinstaller doesn't workw with anything later. If you want to run just the python, using a newer versino should work fine.
```bash
pip3 install —upgrade pip
pip3 install xlrd
pip3 install xlwt
pip3 install pyexcel
pip3 install pyexcel_xlsx
pip3 install pyinstaller
```

For the Mac, there is an Apple Automator app that can be used do drag and drop the .zip file onto which executes the script. You will need to modify the automator app to specify the installation directory.

**Installation Directory Structure:**

-bin: contains the pyinstller executables
-src: contains the Python sources
-config: contains the spreadsheets used to configure NMRAMembers (more on this below)
-build: directory used to run pyinstaller
-work: created by NMRAMembers to unzip the nmra_zip_file and contains the processed files as the program runs
-release: created by NMRAMembers to store the processed .zip output files

THe NMRA files are sent with a name that matches the following pattern:

RRMMMYYYY.zip

-RR: the two-digit numerical region ID
-MMM: the three-letter month (Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec)
-YYYY: the four-digit year

This 9-character name is used as the directory name under the 'work' and 'release' directories to organize the files.

**Building NMRAMembers:**

Pyinstaller has a very finicky build process that has been captured in a shell script.

```bash
cd installation_directory/build
./build.sh
```

**Command line execution:**

```bash
cd installation_directory
./bin/MacOSX/NMRAMembers --help
```

**Help Output:**

NMRA Roster ZIP File Processing Program  <br />
by Erich Whitney  <br />
Copyright (c) 2019, 2020 HomeBrew Engineering  <br />
Version v0.5  <br />
  <br />
usage: NMRAMembers [-h] [-r reassign] [-m, map_file] [-e, email_file] [-w work_dir] [-d dist_dir] [-l] [-s] [-f] nmra_zip_file  <br />
  
This program processes the NMRA roster files from the NMRA .zip file and outputs a directory with the roster .zip files for each division and region found with the members reassigned to their desired divsions as specified in the division reassignment file (-r option) Since the NMRA roster files use numerical identifiers for each region and division, this script uses a mapping file (-m option)  <br />
  <br />
positional arguments:  <br />
&nbsp;&nbsp;nmra_zip_file&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The NMRA ZIP file containing the monthly roster to be processed  <br />
  <br />
optional arguments:  <br />
&nbsp;&nbsp;-h, --help <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Show this help message and exit <br />
&nbsp;&nbsp;-r reassign, --reassign reassign  <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Filename of the .xlsx reassignment file (default: ['./config/NMRA_Division_Reassignments.xlsx'])  <br />
&nbsp;&nbsp;-m, map_file, --map_file map_file  <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Filename of the .xlsx division map file (default: ['./config/NMRA_Region_Division_Map.xlsx'])  <br />
&nbsp;&nbsp;-e, email_file, --email_file email_file  <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Filename of the .xlsx email distribution file (default: ['./config/NMRA_Email_Distribution_List.xlsx'])  <br />
&nbsp;&nbsp;-w work_dir, --work_dir work_dir  <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Name of the work directory (default: ['work'])  <br />
&nbsp;&nbsp;-d dist_dir, --dist_dir dist_dir  <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Name of the directory where all of the final output files go (default: ['release'])  <br />
&nbsp;&nbsp;-l, --long_dir_names
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Use full-length names for the region/division directories instead of a shorter version (default: False)  <br />
&nbsp;&nbsp;-s, --send_emai
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Send out the emails according to the distribution list (default: False)  <br />
&nbsp;&nbsp;-f, --force_override_email  <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Force the override of the email address in the config file in place of the NMRA email address (default: False)  <br />
 <br />
## Documentation
The nmra_zip_file is the .zip file that is sent to the region each month which contains the current membership roster information. NMRAMembers will unzip this file and process the contents.

### Configuration Files

- NMRA_Region_Division_map.xlsx (the --map_file option): This file takes the NMRA Region and Division ID numbers and converts them to strings that are more human-readible. This file MUST be maintained in sync with the current NMRA Region/Division assignments. Since this informatino was taken from the NMRA web site https://www.nmra.org/regions
- NMRA_Division_Reassignments.xlsx (the --reassign option): This file is used to move NMRA members between divisions in their region. 
- NMRA_Email_Distribution_List.xlsx (the --email_file option): This file is used to maintain an email distribution list for the region. 
#### Templates
Templates for the Division Reassignments and the Email Distribution List are provided and these files MUST be copied to the above files in the config directory and modified prior to using NMRAMembers.

### Options
The current version of NMRAMembers does not send emails by default. You can add the --send_email option to do so.

### Reassigning NMRA members from one division to another
This is a policy that is left up to each region so the NMRA National file maintains each member in their assigned division but if they choose to request a transfer within their region, this file allows the region to implement that policy.

### Sending Emails
The email distribution spreadsheet contains a list of each contact within the region that is allowed to see this membership infomration and the contents of this file needs to be maintained by the regional office manager. 
This file provides a column to enter an email address for each recipient, however, by default this address won't be used without adding the --force_email_override option. 
This is because the script is checking to make sure that each recipient is an NMRA member who's listed in the NMRA roster.

The configuration file provides the following options for the 'file' specification:
1) DIVISION: specifies that the recipient should receive their division membership roster .zip file.
2) REGION: specifies that the recipient should receive their region membershiop roster .zip file.
3) NMRA: specifies that the recipient should receive the processed version of the entire NMRA roster .zip file provided.
4) filename.zip: specifies that the recipient should recieve the specified filename.zip from the release directory.
