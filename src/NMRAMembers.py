#!/usr/bin/env python3
###############################################################################
###############################################################################
#
# NMRAMembers.py
# by Erich Whitney
# Copyright (c) 2019-202, BlackCat Engineering
# Version 2.11
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
# v0.8		Fixed bugs in zip file creation
# v0.9		Updated to NMRA 2023 database format
# v1.0		Initial working release
# v2.0		Updated to use Python Pandas in place of older Pyexcel/xlrd/xlwt
#			legacy libraries.
# v2.1		Changed FILE to REGIONFILE cateory in the email distibution file.
#			Fixed bugs with zip codes and dates.
#			Changed over to using .CSV files for all output files
#			Added the NER Coupler mailing list and other lists for the
#			Coupler editor as part of the NER Office Manager job.
#			Added a master XML configuration file to drive how the monthly
#			reports are processed.
#			Currently, this version has not been compiled to binary so it has
#			only been run directly from Python3.
# v2.11		Fixed bugs with zipcode and birthyear numerical format issues.
#			Added AppleScript for email distribution.
###############################################################################
###############################################################################
import sys
import os
import argparse
import re
import glob
import string
import gc
import shutil
#
# These are for manipulating the NMRA spreadsheet reports in .xls/.xlsx format
#
import xlrd
import pyexcel
#
# These are for manipulating the config files in .xlsx/.csv format
#
import pandas as pd
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
import NewsletterFile
import CopyFile
import RosterFile
import EmailDistribution
import NMRAMembersConfig
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
def zip_directory(filename, directory, ziponly=False):
	try:
		os.access(directory, os.R_OK)
	except:
		raise
	pass
#	print('Opening zip file %s for writing from directory %s...' % (filename, directory))
	with zipfile.ZipFile(filename, "w") as zip_ref:
		if (ziponly):
#			print('Adding zip files from directory: %s:' % (os.getcwd()))
			for fname in glob.glob("*.zip"):
#				print('\t%s' % fname)
				zip_ref.write(fname)
			pass
		else:
#			print('Adding roster files from directory: %s:' % (os.getcwd()))
			for dirname, subdirs, files in os.walk(directory):
				zip_ref.write(dirname)
				for fname in files:
#					print('\t%s/%s' % (dirname, fname))
					zip_ref.write(os.path.join(dirname, fname))
				pass
			pass
		pass
	pass
#	print('\n')
pass
#------------------------------------------------------------------------------
#
# Legacy roster files are .xls format, newer roster files are .xlsx format
#
# NOTE: Warnings about OLE2 and file sizes are expected--these are caused by 
# the really old version of Excel these NMRA files are saved in...
#------------------------------------------------------------------------------
def convert_legacy_roster_files(roster_file_dir):
	print("Converting older Excel files to XLSX format for processing...")
	for roster_file in glob.glob("%s/*.xls" % (roster_file_dir)):
		(xlsbase, xlsext) = os.path.splitext(roster_file)
		xlsxfile = "%s.%s" % (xlsbase, input_format)
		print("\tConverting %s to %s...\n" % (roster_file, xlsxfile))
		wb = xlrd.open_workbook(filename=file_name, encoding_override="cp1252", formatting_info = True)
		sh = wb.sheet_by_index(0)
		sheet = pyexcel.Sheet(name=sh.name)
		zipcol=0
		for r in sh.row_range():
			row_data = []
			for c in sh.column_range():
				if (r > 0):
					cell_type = sh.cell_type(r,c)
					if (cell_type == 1):
						year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sh.cell_value(r,c),0)
						cell_date = "%02d/%02d/%04d" % (month, day, year)
					elif (c == zipcol):
						cell_data = "%05d" % (int(sh.cell_value(r,c)))
					else:
						cell_data = sh.cell_value(r,c)
					pass
				else:
					if (zipcol == 0):
						if (sh.cell_value(r,c) == "zip"):
							zipcol = c
						pass
					pass
				pass
			row_data.append(cell_data)
			pass
			sheet.row += row_data
		pass
		sheet.save_as(xlsxfile)
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
	program_version = "v2.12"
	default_config_file=['./config/NMRAMembersConfig.xml']
	default_reassignment_file=['./config/NMRA_Division_Reassignments.xlsx']
	default_map_file=['./config/NMRA_Region_Division_Map.xlsx']
	default_email_file=['./config/NMRA_Email_Distribution_List.xlsx']
	default_work_dir=['work']
	default_dist_dir=['release']
	default_seasonal_members_file=['./RegionSeasonalMembers.xlsx']
	default_distribute=False
	default_override_email=False
	default_legacy_mode=False

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
		'-c', '--config',
		metavar='config',
		nargs=1,
		default=default_config_file,
		required=False,
		help="Filename of the XML configuration file (default: %s)" % (default_config_file)
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
		'-n', '--seasonal',
		metavar='seasonal',
		nargs=1,
		default=default_seasonal_members_file,
		required=False,
		help="Filename of the .xlsx seasonal members file (default: %s)" % (default_seasonal_members_file)
		)
	parser.add_argument(
		'-s', '--send_email',
		action="store_true",
		default=default_distribute,
		required=False,
		help="Send out the emails according to the distribution list (default: %s)" % (default_distribute)
		)
	parser.add_argument(
		'-f', '--force_override_email',
		action="store_true",
		default=default_override_email,
		required=False,
		help="Force the override of the email address in the config file in place of the NMRA email address (default: %s)" % (default_override_email)
		)
	parser.add_argument(
		'-g', '--legacy',
		action="store_true",
		default=default_legacy_mode,
		required=False,
		help="Process older legacy roster files (default: %s)" % (default_legacy_mode)
		)
#------------------------------------------------------------------------------
#
# Start
#
#------------------------------------------------------------------------------
	print("\nNMRA Roster ZIP/Spreadsheet File Processing Program")
	print("by Erich Whitney")
	print("Copyright (c) 2019-2024 BlackCat Engineering")
	print("Version %s" % program_version)
	print("")
	#
	# Handle arguments
	#
	args = parser.parse_args(namespace=myargs)
	zip_file = myargs.nmra_zip_file[0]
	zip_temp = os.path.splitext(os.path.basename(zip_file))
	zip_name = zip_temp[0]
	
	work_dir = myargs.work_dir[0]
	dist_dir = "%s/%s" % (myargs.dist_dir[0], zip_name)
	config_file = myargs.config[0]
	reassign_file = myargs.reassign[0]
	seasonal_file = myargs.seasonal[0]
	map_file = myargs.map_file[0]
	unzip_dir = "%s/%s" % (myargs.work_dir[0], zip_name)
	email_file = myargs.email_file[0]
	send_email = myargs.send_email
	force_override = myargs.force_override_email
	legacy_mode = myargs.legacy
	#
	# Read the XML configuration file
	#
	mode = 'nmra-xlsx'
	if (legacy_mode):
		mode = 'nmra-xls'
	pass
	config = NMRAMembersConfig.NMRAMembersConfig(config_file, mode)
	action_list = config.get_action_list()
	input_format = config.get_input_format()
	output_format = config.get_output_format()
	#
	# process map file
	#
	nmra_map = DivisionMap.DivisionMap()
	nmra_map.read_file(map_file)
	print("---------------------------------------------------------")
	#
	# process reassignment file
	#
	reassignments = MemberReassignments.MemberReassignments()
	reassignments.read_file(reassign_file)
	print("---------------------------------------------------------")
	#
	# process email distribution file
	#
	email_distribution = EmailDistribution.EmailDistribution()
	email_distribution.read_file(email_file)
	print("---------------------------------------------------------")
	#
	# make sure all of the output directories exists
	#
	os.makedirs("%s" % work_dir, exist_ok=True)
	os.makedirs("%s" % dist_dir, exist_ok=True)
	#---------------------------------------------------------------------------------------------
	#
	# expand the NMRA zip file into the working directory
	#
	#---------------------------------------------------------------------------------------------
	print("\nExpanding the NMRA ZIP file...")
	expand_zip_file(zip_file, work_dir)
	(zip_filename, zip_ext) = os.path.splitext(os.path.basename(zip_file))
	roster_file_dir = "%s/%s" % (work_dir, zip_name)
	print("---------------------------------------------------------")
	#---------------------------------------------------------------------------------------------
	#
	# Check the unzipped files and make sure all if the input files exist.
	# Also, determine the region from the unzipped filenames.
	#
	#---------------------------------------------------------------------------------------------
	re1 = re.compile(r'^(\d+)_(.*)') # parses the regionID_filename from a filename
	region = 0
	region_ok = False
	for xfile in glob.glob("%s/*.%s" % (roster_file_dir, input_format)):
		xfile1 = os.path.basename(xfile)
		result = re1.split(xfile1)
		if (len(result) > 1):
			reg = int(result[1])
			if (region == 0):
				region = reg
				region_ok = True
			elif (reg != region):
				region_ok = False
				print("Error: Source file region mismatch (%d <=> %d" % (region, reg))
			pass
		pass
	pass
	file_id = nmra_map.get_file_id(region, 0)
	region_rid = nmra_map.get_region_id(file_id)
	print("Processing NMRA files for the %s Region (NMRA Region ID %d)..." % (region_rid, region))
	#---------------------------------------------------------------------------------------------
	#
	# Copy the RegionSeasonalMembers file into the work directory for processing
	#
	#---------------------------------------------------------------------------------------------
	if (os.path.exists(seasonal_file)):
		seasonal_file_path = "%s/%d_%s" % (unzip_dir, region, os.path.basename(seasonal_file))
		shutil.copyfile(seasonal_file, seasonal_file_path)
	pass
	#---------------------------------------------------------------------------------------------
	#
	# At this point, the NMRA ZIP file has been expanded into the work directory and all of the
	# files are in a subdirectory that was named the same name as the ZIP filename
	#
	#---------------------------------------------------------------------------------------------
	#
	# If it's the file format we need to convert to xlsx format
	#
	if (legacy_mode):
		convert_legacy_roster_files(roster_file_dir)
	pass
	#---------------------------------------------------------------------------------------------
	#
	# At this point if we are in legacy mode, there will be dupicates of all the spreadsheets
	# in the work directory, the original spreadsheets and a copy saved in XLSX format for further
	# processing below...
	#
	#---------------------------------------------------------------------------------------------
	#
	# Create the division and region lists
	#
	parent_dir = os.getcwd()
	#
	# This regular expression parses filenames that come from HQ in the format regionID_filename.ext
	#
	re2 = re.compile(r'^(\d+)_(.+)\.(.+)')
	#
	# Work through all of the file actions
	#
	instance_count = {}	# keep track of how many different times each roster file is used
	#
	# Only process files identified in the XML config file
	#
	#
	# The Newsletter action is for the region newsletter mailing lists, new member and deceased member reports
	#
	all_recipients = {}
	for action in action_list:
		config_files = config.get_files(action)
		if (len(config_files) > 0):
			roster_files = []
			#
			# Map the filename from the XML to the name that should be found in the unzipped archive
			#
			for config_file in config_files:
				config_filename = "%s/%d_%s.%s" % (roster_file_dir, region, config_file, input_format)
				if config_file in instance_count.keys():
					instance_count[config_file] += 1
				else:
					instance_count[config_file] = 1
				pass
				print("\tProcessing %s (%d)..." % (config_filename, instance_count[config_file]))
				roster_file = config_file
				roster_filename = config_filename
				#
				# get a list of the actual files found from the unzipped archive
				#
				files = glob.glob(config_filename)
				#
				# process each file based on the actions for that are defined
				#
				if config_filename in files:
					print("\tPerforming %s Action..." % (action.upper()))
					if (action == 'newsletter'):
						enable_reassignment = False
						print("****Processing Newsletter File: %s, Instance %d" % (roster_filename, instance_count[config_file]))
						nf = NewsletterFile.NewsletterFile(roster_file, instance_count[config_file], unzip_dir, region, config, nmra_map)
						nf.read_file(roster_filename)
						nf_recipients = nf.process(parent_dir)
						all_recipients.update({config_filename : {action : [nf_recipients]}})
						if (not config_filename in all_recipients.keys() ):
							all_recipients.update({config_filename : {action : [nf_recipients]}})
						else:
							current = all_recipients[config_filename]
							all_recipients.update({config_filename : {action : [nf_recipients]}})
							if (not action in current.keys()):
								all_recipients.update({config_filename : {action : [nf_recipients]}})
							else:
								all_recipients[config_filename].get(action).append(nf_recipients)
							pass
						pass
						del nf
						print("---------------------------------------------------------")
					#
					# The Copy action is for any files that just get copied without any member reassignment action
					#
					elif (action == 'copy'):
						enable_reassignment = False
						print("\t\tCopying NMRA Roster file: %s" % (roster_file))
						cp = CopyFile.CopyFile()
						cp_recipients = cp.process(roster_file, instance_count[config_file], roster_filename, unzip_dir, output_format, region, config, nmra_map)
						if (not config_filename in all_recipients.keys()):
							all_recipients.update({config_filename : {action : [cp_recipients]}})
						else:
							current = all_recipients[config_filename]
							if (not action in current.keys()):
								all_recipients.update({config_filename : {action : [nf_recipients]}})
							else:
								all_recipients[config_filename].get(action).append(cp_recipients)
							pass
						pass
						del cp
						print("---------------------------------------------------------")
					#
					# The Reassignment action is for region and division roster files where we have to correct member division assignments based
					# on the reassignments table provided in the NMRE_Division_Reassignments.xlsx spreadsheet
					#
					elif (action == 'reassignment'):
						enable_reassignment = True
						print("\t\tProcessing NMRA Roster file: %s" % (roster_file))
						rf = RosterFile.RosterFile(roster_file, instance_count[config_file], enable_reassignment, unzip_dir, region, config, nmra_map, reassignments, legacy_mode)
						rf.read_file(roster_filename, legacy_mode)
						rf_recipients = rf.process(email_distribution, parent_dir, dist_dir, zip_filename, force_override, legacy_mode)
						if (not config_filename in all_recipients.keys()):
							all_recipients.update({config_filename : {action : [rf_recipients]}})
						else:
							current = all_recipients[config_filename]
							if (not action in current.keys()):
								all_recipients.update({config_filename : {action : [rf_recipients]}})
							else:
								all_recipients[config_filename].get(action).append(rf_recipients)
							pass
						pass
						del rf
						print("---------------------------------------------------------")
					pass
				else:
					print("\tSkipping NMRA Roster file: %s (it wasn't in the NMRA .zip file)" % (roster_file))
				pass
			pass
			gc.collect()
		pass
	pass
	#
	# Copy any region files over to divisions as necessary
	#
	file_list = email_distribution.get_file_list()
	if (len(file_list) > 0):
		print("\nCopying individual region files to divisions per the email distribution list requests...")
		print("-------------------------------------------------------------------------------------------")
	pass
	for entry in file_list:
		for div_fid, filename in entry.items():
			div_name = nmra_map.get_division(div_fid)
			source_path = "%s/%s_Region/%s" % (unzip_dir, region_rid, filename)
			dest_path   = "%s/%s_Region-%s_Division/%s_Region_%s" % (unzip_dir, region_rid, div_name, region_rid, filename)
			print("\tCopying Region File %-30s to: %s" % (filename, dest_path))
			shutil.copyfile(source_path, dest_path)
		pass
	pass
	if (len(file_list) > 0):
		print("-------------------------------------------------------------------------------------------\n")
	pass
	#
	# Now that all of the files are processed in the working directory
	# Zip them up to their respective release output directories
	#
	processed_files=[]
	print("Creating .zip files for each Region/Division...")
	zip_work_dir="%s/%s" % (work_dir, zip_name)
	os.chdir(zip_work_dir)
#	print("Working directory: %s" % (zip_work_dir))
	print("Creating all of the Region and Divsion ZIP files in: %s" % (dist_dir))
	output_files = glob.glob("*")
	for output_file in output_files:
		if (os.path.isdir(output_file)):
#			print("Directory to Zip: %s" % output_file)
			bn = os.path.basename(output_file)
			zip_file_name = "%s/%s/%s.zip" % (parent_dir, dist_dir, bn)
#			print("\t%s/%s" % (dist_dir, os.path.basename(zip_file_name)))
			roster_directory = output_file
			zip_directory(zip_file_name, roster_directory, False)
			processed_files.append(zip_file_name)
		pass
	pass
	print("---------------------------------------------------------")
	#
	# This creates one Zip of all of the files to archive the entire monthly report
	#
	print("Creating a .zip archive of this monthly report...")
	os.chdir(parent_dir)
	release_parent_dir = "%s" % (dist_dir)
	os.chdir(release_parent_dir)
#	print("Release directory: %s" % release_parent_dir)
	full_zip_file_name = "%s/%s/%s_processed.zip" % (parent_dir, myargs.dist_dir[0], zip_filename)
	print("\nArchive of %s is in: %s" % (zip_name, full_zip_file_name))
	zip_directory(full_zip_file_name, zip_name, True)
	print("---------------------------------------------------------")
	#
	# Distribute the zip files to the distribution list
	#
	os.chdir(parent_dir)
	if (send_email):
		print("Sending emails to the distribution list...")
		email_distribution.send_emails(zip_filename)
	else:
		print("Skipping email distribution")
		print("")
		print("Here's the steps that need to be done manually...")
		email_distribution.print_email_list()
	pass
	print("---------------------------------------------------------")
	print("")
	print("\nDone.\n")
pass
#
# and go!
#
if (__name__ == "__main__"):
	main()
pass
