#!/usr/bin/env python3
###############################################################################
###############################################################################
#
# NMRAMembers.py
# by Erich Whitney
# Copyright (c) 2019, HomeBrew Engineering
# Version 0.4
#
# This program is designed to process the monthly NMRA membership reports sent
# to the regions by NMRA National. These reports come in the form of a .zip
# file that contains several Excel spreadsheets that are saved in a very old
# .xls format (Excel 5.0/95). This poses some problems because the files are
# difficult to process using most Excel scripting options. For this reason,
# the xlrd/xlwt libraries are used because they are able to handle these files.
#
# The NMRA National sends these files to the region which is then supposed to
# split the information out for each division. However, due to the fact that
# each division's information is spread across multiple spreadsheets, the
# process of managing this data is cumbersome at best. So this program was
# created to automate that part of the job.
#
# The core function of this program is to start with the NMRA's .zip file and
# expand it into the set of monthly report files. It then processes each of
# those files, writing out a separate file for each division. It also creates
# a new region version that contains just the data for that region--any other
# region information is broken out so it's easier to tell if there's either
# incoming or outgoing member information.
#
# Finally, as part of the process of creating these report files, this program
# can optionally implement a division-level member reassignment operation.
# This is because any NMRA member who wishes to belong to another division
# within their region, needs to make that request to the region who implements
# that request. Again, due to the cumbersome nature of these spreadsheets, an
# automated solution as performed by this program is a lot more sensible.
#
# This program uses two configuration files in addition to the NMRA monthly
# .zip input file. The NMRA Region/Division Map file contains the current
# association between any region and division and its 4-digit identifier.
# This is useful because the NMRA report files use numbers and humans much
# prefer to see textual descriptions which are more familiar. The other
# configuration file is the NMRA Region/Division reassignment file. This is
# simply a list of each NMRA member who wishes to change their division
# affiliation. The members NMRA number, new region/division, and first and last
# names are in this file. The first and last names are used to verify that
# the person who's NMRA number is used in the list matches who the NMRA thinks
# they are.
#
# A final note about this version of the program. It is written in Python 3.8
# and makes extensive use of Python Classes. The code was written with the idea
# that it could be adapted over time as needs change and be able to handle any
# change in the Excel data format should that ever change.
#
# Each Excel report file has basically the same format. The first line contains
# the column headings and every row after that represents one NMRA member.
# Because these column headings can change from file to file, this program only
# looks for a minimum subset of the columns it needs to make decisions, namely,
# the NMRA Member ID, Region, Division, First Name, and Last Name. All other
# column headings are just copied over and the column order is preserved as is
# the data. Only the data is copied over--no formatting is preserved so manual
# editing of the Excel files prior to running the program is not recommended.
#
# xlrd/xlwt: These are currently the libraries that do the work of handing the
# NMRA spreadsheet data. xlrd can only read .xls and xlwt can only write .xls
# so this program reads each report spreadsheet in one at a time, processes it,
# and then creates N copies for each output region and division as necessary.
# Since .xls is pretty universally exchanged between spreadsheet programs, this
# file format is preserved. There are utilities included in the ./bin directory
# that let you convert between .csv, .xls, and .xlsx.
#
###############################################################################
#
# Change Log:
# v0.1		Initial Version
# v0.2		Added the from/to region/division to the reassignment file and
#			added the check to make sure the from and to regions are the same
#			because changing regions with this program is not allowed.
# v0.3		Broke up the classes into individual files for easier code reuse
# v0.4		Addressed more Excel dates showing up in the wrong format in the
#			output files.
###############################################################################
###############################################################################
import sys
import os
import argparse
import re
import glob
import string
#
# These are for manipulating the NMRA spreadsheet reports in .xls format
#
#import xlrd
import xlwt
#
# These are for manipulating the config files in .xlsx format
#
import pyexcel
import pyexcel_xlsx
#
# Used to manage the NMRA zip file and to package up the output files
#
import zipfile
###############################################################################
###############################################################################
#
# Program Classes
#
###############################################################################
import DivisionMap
import MemberReassignments
import MemberInfo
import RosterFile
import EmailDistribution
#------------------------------------------------------------------------------
#
# Define the namespace for arguments in parse_args
#
#------------------------------------------------------------------------------
class myargs:
	pass
###############################################################################
###############################################################################
#
# Helper Functions
#
###############################################################################
###############################################################################
#------------------------------------------------------------------------------
#
# Expand the given zip file to the given directory
#
#------------------------------------------------------------------------------
def expand_zip_file(filename, directory):
	if (not os.access(filename, os.R_OK)):
		print('Error: Zip File %s is not readible!' % filename)
		print('')
		sys.exit(-1)
	else:
		with zipfile.ZipFile(filename, "r") as zip_ref:
			zip_ref.extractall(directory)
		pass
	pass
pass
#------------------------------------------------------------------------------
#
# Zip up the contents of the given directory to the given zip file
#
#------------------------------------------------------------------------------
def zip_directory(filename, directory):
	try:
		os.access(directory, os.R_OK)
	except:
		raise
	pass
	with zipfile.ZipFile(filename, "w") as zip_ref:
		for dirname, subdirs, files in os.walk(directory):
			zip_ref.write(dirname)
			for fname in files:
				zip_ref.write(os.path.join(dirname, fname))
			pass
		pass
	pass
pass
###############################################################################
###############################################################################
#
# Start of main procedure
#
###############################################################################
###############################################################################
def main():
#------------------------------------------------------------------------------
#
# Create the program argument definitions
#
#------------------------------------------------------------------------------
	program_version = "v0.5"
	default_reassignment_file=['./config/NMRA_Division_Reassignments.xlsx']
	default_map_file=['./config/NMRA_Region_Division_Map.xlsx']
	default_email_file=['./config/NMRA_Email_Distribution_List.xlsx']
	default_work_dir=['./work']
	default_dist_dir=['./release']
	default_distribute=False
	default_long=False

	parser = argparse.ArgumentParser(
		description='This program processes the NMRA roster files from the NMRA .zip file and outputs a directory with the roster .zip files for each division and region found with the members reassigned to their desired divsions as specified in the division reassignment file (-r option)\nSince the NMRA roster files use numerical identifiers for each region and division, this script uses a mapping file (-m option)'
		)
	parser.add_argument(
		'nmra_zip_file',
		metavar='nmra_zip_file',
		nargs=1,
		help='The NMRA ZIP file containing the monthly roster to be processed'
		)
	parser.add_argument(
		'-r', '--reassign',
		metavar='reassign',
		nargs=1,
		default=default_reassignment_file,
		required=False,
		help="Filename of the .xlsx reassignment file (default: %s)" % (default_reassignment_file)
		)
	parser.add_argument(
		'-m,', '--map_file',
		metavar='map_file',
		nargs=1,
		default=default_map_file,
		required=False,
		help="Filename of the .xlsx division map file (default: %s)" % (default_map_file)
		)
	parser.add_argument(
		'-e,', '--email_file',
		metavar='email_file',
		nargs=1,
		default=default_email_file,
		required=False,
		help="Filename of the .xlsx email distribution file (default: %s)" % (default_email_file)
		)
	parser.add_argument(
		'-w', '--work_dir',
		metavar='work_dir',
		nargs=1,
		default=default_work_dir,
		required=False,
		help="Name of the work directory (default: %s)" % (default_work_dir)
		)
	parser.add_argument(
		'-d', '--dist_dir',
		metavar='dist_dir',
		nargs=1,
		default=default_dist_dir,
		required=False,
		help="Name of the directory where all of the final output files go (default: %s)" % (default_dist_dir)
		)
	parser.add_argument(
		'-l', '--long_dir_names',
		action="store_true",
		default=default_long,
		required=False,
		help="Use full-length names for the region/division directories instead of a shorter version (default: %s)" % (default_long)
		)
	parser.add_argument(
		'-s', '--send_email',
		action="store_true",
		default=default_distribute,
		required=False,
		help="Send out the emails according to the distribution list (default: %s)" % (default_distribute)
		)
#------------------------------------------------------------------------------
#
# Start
#
#------------------------------------------------------------------------------
	print("\nNMRA Roster ZIP File Processing Program")
	print("by Erich Whitney")
	print("Copyright (c) 2019, 2020 HomeBrew Engineering")
	print("Version %s" % program_version)
	print("")
	#
	# Handle arguments
	#
	args = parser.parse_args(namespace=myargs)
	zip_file = myargs.nmra_zip_file[0]
	zip_temp = os.path.splitext(os.path.basename(zip_file))
	zip_name = zip_temp[0]
	use_long = myargs.long_dir_names
	
	work_dir = myargs.work_dir[0]
	dist_dir = "%s/%s" % (myargs.dist_dir[0], zip_name)
	reassign_file = myargs.reassign[0]
	map_file = myargs.map_file[0]
	unzip_dir = "%s/%s" % (myargs.work_dir[0], zip_name)
	email_file = myargs.email_file[0]
	send_email = myargs.send_email
	#
	# process map file
	#
	div_map = DivisionMap.DivisionMap()
	div_map.read_file(map_file)
	#
	# process reassignment file
	#
	reassignments = MemberReassignments.MemberReassignments()
	reassignments.read_file(reassign_file)
	#
	# process email distribution file
	#
	email_distribution = EmailDistribution.EmailDistribution()
	email_distribution.read_file(email_file)
	#
	# make sure all of the output directories exists
	#
	os.makedirs("%s" % work_dir, exist_ok=True)
	os.makedirs("%s" % dist_dir, exist_ok=True)
	#
	# expand the NMRA zip file into the working directory
	#
	print("\nExpanding the NMRA ZIP file...")
	expand_zip_file(zip_file, work_dir)
	zip_filename = os.path.splitext(os.path.basename(zip_file))
	roster_file_dir = "%s/%s" % (work_dir, zip_name)
	
	#
	# Create the division and region lists
	#
	print("\nProcessing Roster Files...")
	print("NOTE: Warnings about OLE2 and file sizes are expected--these are caused by the really old version of Excel these NMRA files are saved in...")
	roster_files = glob.glob("%s/*.xls" % (roster_file_dir))
	for roster_file in roster_files:
		print("\tProcessing NMRA Roster file: %s" % roster_file)
		rf = RosterFile.RosterFile(unzip_dir, div_map, use_long, reassignments)
		rf.read_file(roster_file)
		rf.process(email_distribution)
	pass
	#
	# Zip up the output directories
	#
	parent_dir = os.getcwd()
	os.chdir(work_dir)
	print("\nCreating all of the Region and Divsion ZIP files in: %s" % (dist_dir))
	output_files = glob.glob("%s/*" % zip_name)
	for output_file in output_files:
		if (os.path.isdir(output_file)):
			bn = os.path.basename(output_file)
			zip_file_name = "%s/%s/%s.zip" % (parent_dir, dist_dir, bn)
			print("\t%s/%s" % (dist_dir, os.path.basename(zip_file_name)))
			roster_directory = output_file
			zip_directory(zip_file_name, roster_directory)
		pass
	pass
	#
	# Email the zip files to the distribution list
	#
	os.chdir(parent_dir)
	if (send_email):
		print("Sending emails to the distribution list...")
		dist_path = "%s/%s" % (parent_dir, dist_dir)
		email_distribution.send_emails(zip_name, dist_path, myargs.nmra_zip_file[0])
	else:
		print("Skipping email distribution")
	pass

	print("\nDone.\n")
pass
#
# and go!
#
if (__name__ == "__main__"):
	main()
pass
