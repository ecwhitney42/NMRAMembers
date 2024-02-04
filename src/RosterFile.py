#------------------------------------------------------------------------------
#
# Class RosterFile
#
# Process a roster file
#
# This class manages the reading and writing of the individual
# roster files found after unzipping the given NMRA zip file.
#
# Each file is read in and if any of its members are in the
# reassignment file, they are moved to the new region/division.
#
# The output of each file is placed in a directory in the format
# region/divsion_filename.zip. This is so in the end, each divsiion
# can get a zip file of just its data with no other division
# information. A full copy of all the region data is saved in
# a zip file of the format region_filename.zip
#
# Since each month the NMRA send a zip file with the date in it,
# the output is saved in a subdirectory with the same name so
# that each month does not overwrite a previous month.
#------------------------------------------------------------------------------
import sys
import pandas as pd 
import DivisionMap
import xlrd
import xlwt
import os
import re
import warnings

class RosterFile:
	#
	# Constructor
	#
	
	def __init__(self, roster_file, instance, enable_reassignment, work_dir, region, config, nmra_map, reassignments, legacy_mode):
		self.roster_file = roster_file
		self.region = region
		self.work_dir = work_dir
		self.nmra_map = nmra_map
		self.reassignments = reassignments
		self.legacy_mode = legacy_mode
		self.region_filenames = {}
		self.regions = []
		self.division_filenames = {}
		self.divisions = []
		self.editor_filenames = {}
		self.editor = []
		self.report_name = ""
		self.roster_rxf = None
		self.roster_rdf = None
		self.roster_wrdf = {}
		self.roster_wddf = {}
		self.roster_wedf = {}
		self.roster_nrows = 0
		self.roster_ncols = 0
		self.this_region = region
		self.region_header = 'region'
		self.enable_reassignment = enable_reassignment
		self.instance = instance

		self.validate = config.get_validate(roster_file)
		self.roster_ifmt = config.get_input_format()
		self.roster_ofmt = config.get_output_format()
		self.date_fields = config.get_date_fields('reassignment', roster_file, self.instance)
		self.recipient_list = config.get_recipients('reassignment', roster_file, self.instance)
		self.date_format = config.get_date_format()
		self.recipients = {}
		for recipient in self.recipient_list:
			self.recipients.update({recipient : []})
		pass 
		
#		print("The recipients are: %s" % (self.recipient_list))
#		print("Date Fields are: %s" % (self.date_fields))
#		print("Date Format is: %s" % (self.date_format))

	pass

	def get_nmra_id(self, row):
		return '%s' % (self.roster_rdf.at[row, 'id'])
	pass
	def get_lname(self, row):
		lname = '%s' % (self.roster_rdf.at[row, 'lname'])
		' '.join(lname.split())		
		return '%s' % (lname)
	pass
	def get_fname(self, row):
		fname = '%s' % (self.roster_rdf.at[row, 'fname'])
		' '.join(fname.split())		
		return '%s' % (fname)
	pass
	def get_email(self, row):
		if 'email' in self.roster_rdf.columns:
			a_email	= '%s' % (self.roster_rdf.at[row, 'email'])
		else:
			a_email = ''
		pass
		return a_email
	pass
	def get_division(self, row):
		t_division	= self.roster_rdf.at[row, 'division']
		try:
			a_division	=    int(t_division)
		except ValueError:
			a_division  = 0
		pass
		return a_division
	pass
	def get_region(self, row):
		t_region	= self.roster_rdf.at[row, self.region_header]
		try:
			a_region	=    int(t_region)
		except ValueError:
			a_region    = 0
		pass
		return a_region
	pass

	#
	# this method returns true if the given member ID is in the list of reassignments 
	# and the first and last name matches and the from division is not the to division 
	#
	def is_member_reassigned(self, nmra_id, row):
		ret_val = False
		
		reassign_this_member	= self.reassignments.has_member(nmra_id)
		
		if (reassign_this_member):
			member_div			= self.get_division(row)
			member_id			= self.get_nmra_id(row)
			member_lname		= self.get_lname(row)
			member_fname		= self.get_fname(row)
			reassign_from_div	= self.reassignments.get_from_division(nmra_id)
			reassign_to_div		= self.reassignments.get_to_division(nmra_id)
			reassign_id			= self.reassignments.get_nmra_id(nmra_id)
			reassign_lname		= self.reassignments.get_lname(nmra_id)
			reassign_fname		= self.reassignments.get_fname(nmra_id)
			if ((member_div == reassign_from_div) and (member_div != reassign_to_div)):
#				if ((member_lname == reassign_lname) and (member_fname == reassign_fname)):
				if ((member_id == reassign_id)):
					ret_val = True
				else:
#					print("WARNING: Reassignment found for member %s, however, the name doesn't match so no reassignment" % nmra_id)
					print("WARNING: Reassignment found for member %s, however, the member ID doesn't match so no reassignment" % nmra_id)
					print("From Division: '%d' <=> '%d'" % (member_div, reassign_from_div))
					print("  To Division: '%d' <=> '%d'" % (member_div, reassign_to_div))
					print("      NMRA ID: '%s' <=> '%s'" % (member_id, reassign_id))
					print("   Last Name:  '%s' <=> '%s'" % (member_lname, reassign_lname))
					print("  First Name:  '%s' <=> '%s'" % (member_fname, reassign_fname))
						  
				pass
			pass
		pass
		return ret_val
	pass
	#
	# Read the given roster file and save the information necessary to create a new workbook from it.
	# This is where we are able to read in the crusty old version of Excel files.
	#
	def read_file(self, filename, legacy_mode):
		#
		# Get the sheet
		#
#		self.roster_exf = pd.ExcelFile(filename)
#		self.roster_rdf = self.roster_exf.parse(parse_dates=self.date_fields)
		self.roster_rdf = pd.read_excel(filename, date_format=self.date_format, dtype='string') #parse_dates=self.date_fields)
		#
		# Make all of the column headings lower case
		#
#		print("%s" % (self.roster_rdf.columns))	
		for col in range(0, len(self.roster_rdf.columns)):
			old = self.roster_rdf.columns[col]
			new = old.lower()
			if (new == 'memberregion'):
				self.region_header = new
			pass
#			print("\t%s=>%s" % (old, new))
			self.roster_rdf.rename(columns={old : new}, inplace=True)
			if ((new == 'birthyear') or (new == 'zip')):
#			if (new == 'zip'):
				self.roster_rdf[new] = self.roster_rdf[new].astype('string')
			pass
		pass
		#
		# Fix the date field
		#
		for field in self.date_fields:
#		print("Setting column %s to %s format..." % (field, self.date_format))
			self.roster_rdf[field]=pd.to_datetime(self.roster_rdf[field])
			self.roster_rdf[field]=self.roster_rdf[field].dt.strftime(self.date_format)
		pass
		#
		# take out the extra numbers in the filename and save the root string
		# to a report_name that way we can write out cleaner filenames
		#
		report_file = os.path.splitext(os.path.basename(filename))
		self.report_name = report_file[0]
		re1 = re.compile(r'^(\d+)_(.*)')
		self.report_name = re1.sub(r'\2', self.report_name)
		re2 = re.compile(r'(.*)(\d+)$')
		self.report_name = re2.sub(r'\1', self.report_name)
	pass
	#
	# Create a standard filename
	#
	def format_filename(self, dir_path):
		filename = "%s/%s.%s" % (dir_path, self.report_name, self.roster_ofmt)
		return(filename)
	pass
	#
	# Create a workbook for a given output file.
	#
	def create_workbook(self, nmra_id, dir_path, recipient):
		os.makedirs("%s" % dir_path, exist_ok=True)
		filename = self.format_filename(dir_path)
		if (recipient == 'region'):
			if (not nmra_id in self.region_filenames.keys()):
#				print("**Creating Region Workbook %s, %s" % (nmra_id, filename))
				self.region_filenames[nmra_id] = filename
				self.roster_wrdf[nmra_id] = pd.DataFrame(columns=self.roster_rdf.columns)
			pass
		elif (recipient == 'division'):
			if (not nmra_id in self.division_filenames.keys()):
#				print("**Creating Division Workbook %s, %s" % (nmra_id, filename))
				self.division_filenames[nmra_id] = filename
				self.roster_wddf[nmra_id] = pd.DataFrame(columns=self.roster_rdf.columns)
		elif (recipient == 'editor'):
			if (not nmra_id in self.editor_filenames.keys()):
#				print("**Creating Editor Workbook %s, %s" % (nmra_id, filename))
				self.editor_filenames[nmra_id] = filename
				self.roster_wedf[nmra_id] = pd.DataFrame(columns=self.roster_rdf.columns)
			pass
		pass
		#
		# Create a new workbook to copy it into for processing, add the header
		#
		
	pass
	#
	# Copy member info from current worksheet to given output worksheet
	#
	def write_row(self, file_id, row, recipient):
		warnings.simplefilter(action="ignore", category=FutureWarning)
		if (recipient == 'region'):
			wdr_row = len(self.roster_wrdf[file_id].index)
#			print("1:Writing new row %d from row %d for id: %s" % (wdr_row, row, file_id))
			for col in self.roster_rdf.columns:
				self.roster_wrdf[file_id].loc[wdr_row, col] = self.roster_rdf.at[row, col]
			pass
		elif (recipient == 'division'):
			wdr_row = len(self.roster_wddf[file_id].index)
#			print("2:Writing new row %d from row %d for id: %s" % (wdr_row, row, file_id))
			for col in self.roster_rdf.columns:
				self.roster_wddf[file_id].loc[wdr_row, col] = self.roster_rdf.at[row, col]
			pass
		elif (recipient == 'editor'):
			wdr_row = len(self.roster_wedf[file_id].index)
#			print("3:Writing new row %d from row %d for id: %s" % (wdr_row, row, file_id))
			for col in self.roster_rdf.columns:
				self.roster_wedf[file_id].loc[wdr_row, col] = self.roster_rdf.at[row, col]
			pass
		pass
	pass
	#
	# Copy member info from current worksheet to given output worksheet
	# and reassign this member to the given region/division if necessary
	#
	def reassign_member(self, file_id, row, new_region, new_division, recipient):
		if (recipient == 'region'):
			wdr_row = len(self.roster_wrdf[file_id].index)
#			print("4:Writing new row %d from row %d for id: %s" % (wdr_row, row, file_id))
			for col in self.roster_rdf.columns:
				if (col == self.region_header):
					self.roster_wrdf[file_id].loc[wdr_row, col] = new_region
				elif (col == 'division'):
					self.roster_wrdf[file_id].loc[wdr_row, col] = new_division
				else:
					self.roster_wrdf[file_id].loc[wdr_row, col] = self.roster_rdf.at[row, col]
				pass
			pass
		elif (recipient == 'division'):
			wdr_row = len(self.roster_wddf[file_id].index)
#			print("5:Writing new row %d from row %d for id: %s" % (wdr_row, row, file_id))
			for col in self.roster_rdf.columns:
				if (col == self.region_header):
					self.roster_wddf[file_id].loc[wdr_row, col] = new_region
				elif (col == 'division'):
					self.roster_wddf[file_id].loc[wdr_row, col] = new_division
				else:
					self.roster_wddf[file_id].loc[wdr_row, col] = self.roster_rdf.at[row, col]
				pass
			pass
		elif (recipient == 'editor'):
			wdr_row = len(self.roster_wedf[file_id].index)
#			print("6:Writing new row %d from row %d for id: %s" % (wdr_row, row, file_id))
			for col in self.roster_rdf.columns:
				if (col == self.region_header):
					self.roster_wedf[file_id].loc[wdr_row, col] = new_region
				elif (col == 'division'):
					self.roster_wedf[file_id].loc[wdr_row, col] = new_division
				else:
					self.roster_wedf[file_id].loc[wdr_row, col] = self.roster_rdf.at[row, col]
				pass
			pass
		pass
	pass
	#
	# Write a member to the given row and check if they are reassigned
	#
	def write_member(self, nmra_id, row, dataframe_id, recipient):
		m_region = self.get_region(row)
		if (m_region != self.region):
			self.write_row(dataframe_id, row, recipient)
		else:
			m_division = self.get_division(row)
			#
			# check to see if this member is reassigned or not
			#
			reassign = self.is_member_reassigned(nmra_id, row)
			if (reassign):
				r_region = self.reassignments.get_to_region(nmra_id)
				r_division = self.reassignments.get_to_division(nmra_id)
			else:
				r_region = m_region
				r_division = m_division
			pass
			#
			# Write to the division file if this member is in a division
			#
			if (reassign):
				print("\t\tProcessing division reassignment for NMRA member %s from division %02d%02d to division %02d%02d" % (nmra_id, m_region, m_division, r_region, r_division))
				self.reassign_member(dataframe_id, row, r_region, r_division, recipient)
			else:
				self.write_row(dataframe_id, row, recipient)
			pass
		pass
	pass

	#
	# Save the current output workbooks
	#
	def save_workbooks(self):
		for recipient in self.recipient_list:
			if (recipient == 'division'):
				for div_fid in self.division_filenames.keys():
#					print("**Saving Division Workbook %s, %s" % (div_fid, self.division_filenames[div_fid]))
					self.roster_wddf[div_fid].to_csv(self.division_filenames[div_fid], index=False, date_format=self.date_format)#, float_format='{:.0}'.format)
				pass
			elif (recipient == 'region'):
				for reg_fid in self.region_filenames.keys():
#					print("**Saving Region Workbook %s, %s" % (reg_fid, self.region_filenames[reg_fid]))
					self.roster_wrdf[reg_fid].to_csv(self.region_filenames[reg_fid], index=False, date_format=self.date_format)#, float_format='{:.0}'.format)
				pass
			elif (recipient == 'editor'):
				for ed_fid in self.editor_filenames.keys():
#					print("**Saving Region Workbook %s, %s" % (self.editor_filenames[ed_fid]))
					self.roster_wedf[ed_fid].to_csv(self.editor_filenames[ed_fid], index=False, date_format=self.date_format)#, float_format='{:.0}'.format)
				pass
			pass
		pass
	pass
	#
	# Processes the current workbook in memory
	#
	def process(self, distribution, parent_dir, dist_dir, zip_filename, force_override, legacy_mode):
		#
		# Iterate through the rows of each roster worksheet and split into new division and region workbooks
		#
		a_id		= ""
		a_lname		= ""
		a_fname		= ""
		a_division	= 0
		a_region	= 0
		r_division	= 0
		r_region	= 0
		r_lname		= ""
		r_fname		= ""
		r_email		= ""
		reg_fid		= ""
		div_fid		= ""
		
#		print("********Iterating through %d rows of %s for recipients: %s" % (len(self.roster_rdf.index), self.roster_file, self.recipient_list))
#
#			print("********This file will be used to validate the email distribution list members")
#		pass
		for row in self.roster_rdf.index:
			#
			# find the required parameters for this member
			#
			a_id = self.get_nmra_id(row)
			a_lname = self.get_lname(row)
			a_fname = self.get_fname(row)
			a_email = self.get_email(row)
			a_division = self.get_division(row)
			a_region = self.get_region(row)
			reg_fid = self.nmra_map.get_file_id(a_region, 0)
			this_reg_fid = self.nmra_map.get_file_id(self.region, 0)
			this_reg_rid = self.nmra_map.get_region_id(this_reg_fid)
			this_reg_name = self.nmra_map.get_region_id(this_reg_fid)
			div_fid = self.nmra_map.get_file_id(a_region, a_division)
			if (self.enable_reassignment and self.reassignments.has_member(a_id)):
				r_region	= self.reassignments.get_to_region(a_id)
				r_division	= self.reassignments.get_to_division(a_id)
				r_lname		= self.reassignments.get_lname(a_id)
				r_fname		= self.reassignments.get_fname(a_id)
				if (self.is_member_reassigned(a_id, row)):
					reg_fid = self.nmra_map.get_file_id(r_region, 0)
					div_fid = self.nmra_map.get_file_id(r_region, r_division)
				pass
			else:
				r_division	= a_division
				r_region	= a_region
				r_lname		= a_lname
				r_fname		= a_fname
			pass
			if self.nmra_map.has_region_id(reg_fid):
				reg_name = self.nmra_map.get_region_id(reg_fid)
			else:
				raise ValueError("Region ID for %s not found in the map" % reg_fid)
			pass
			#
			# Make sure we have mapped this region/division
			#
			if self.nmra_map.has_region_id(reg_fid):
				#
				# The difference between long and short region names comes from
				# the NMRA Region/Division MAP file that has the RID column
				#
				reg_name = self.nmra_map.get_region_id(reg_fid)
			else:
				raise ValueError("Region ID for %s not found in the map" % reg_fid)
			pass
			if self.nmra_map.has_division(div_fid):
				div_name = self.nmra_map.get_division(div_fid)
			else:
				raise ValueError("Division for %s not found in the map" % div_fid)
			pass
#			print("**********Found member in %s Region, %s Division" % (reg_name, div_name))
			#
			# Create a workbook for each division and region encountered
			#
			# If it's a just a region entry, put it in a region file otherwise break out by division
			#
			# This converts the 4-digit _fid code to a text _name string if it's in the division map
			#
			# members in a file from other regions just go into this regions files.
			#
			region_only = False
			for recipient in self.recipient_list:
#				print("%s" % (recipient))
				if (recipient == 'region'):
					if (not a_region == self.region):
						if (not this_reg_fid in self.regions):
							region_only = True
							self.regions.append(this_reg_fid)
							reg_dir = "%s/%s_Region" % (self.work_dir, this_reg_rid)
							self.create_workbook(this_reg_fid, reg_dir, recipient)
						pass
						if (recipient in self.recipient_list):
							if (not recipient in self.recipients.keys()):
								self.recipients.update({recipient: [self.format_filename(reg_dir)]})
							else:
								self.recipients[recipient].append(self.format_filename(reg_dir))
							pass
						pass
						self.write_member(a_id, row, this_reg_fid, recipient)
					else:
						if ((not reg_fid in self.regions) and (recipient in self.recipient_list)):
							self.regions.append(reg_fid)
							reg_dir = "%s/%s_Region" % (self.work_dir, reg_name)
							self.create_workbook(reg_fid, reg_dir, recipient)
						pass
						if (recipient in self.recipient_list):
							if (not recipient in self.recipients.keys()):
								self.recipients.update({recipient: [self.format_filename(reg_dir)]})
							else:
								self.recipients[recipient].append(self.format_filename(reg_dir))
							pass
						pass
						self.write_member(a_id, row, reg_fid, recipient)
				elif (recipient == 'division'):
#					print("%d <=> %d" % (a_region, self.region))
					if (not (a_region == self.region)):
						if (not reg_fid in self.regions):
							region_only = True
							self.regions.append(this_reg_fid)
							reg_dir = "%s/%s_Region" % (self.work_dir, this_reg_rid)
							self.create_workbook(this_reg_fid, reg_dir, recipient)
						pass
						if (recipient in self.recipient_list):
							if (not recipient in self.recipients.keys()):
								self.recipients.update({recipient: [self.format_filename(reg_dir)]})
							else:
								self.recipients[recipient].append(self.format_filename(reg_dir))
							pass
						pass
						self.write_member(a_id, row, this_reg_fid, recipient)
					else:
						if (r_division == 0):
							print("NMRA-ERROR!: Division Member %s: %-20s %-45s has their division ID set to 0 from the NMRA report, therefore they won't appear in their division report!" % (a_id, "%s %s" % (a_fname, a_lname), "(%s)" % a_email))
						else:
							if ((not div_fid in self.divisions) and (recipient in self.recipient_list)):
								self.divisions.append(div_fid)
								div_dir = "%s/%s_Region-%s_Division" % (self.work_dir, reg_name, div_name)
								self.create_workbook(div_fid, div_dir, recipient)
								if (recipient in self.recipient_list):
									if (not recipient in self.recipients.keys()):
										self.recipients.update({recipient: [self.format_filename(div_dir)]})
									else:
										self.recipients[recipient].append(self.format_filename(div_dir))
									pass
								pass
							pass
							self.write_member(a_id, row, div_fid, recipient)
						pass
					pass
				elif (recipient == 'editor'):
					if ('editor' in self.recipient_list):						
						ed_fid = this_reg_fid
						self.editor.append(ed_fid)
						ed_dir = "%s/%s_Region_Editor" % (self.work_dir, this_reg_rid)
						self.create_workbook(ed_fid, ed_dir, recipient)
						if (not recipient in self.recipients.keys()):
							self.recipients.update({recipient : [self.format_filename(ed_dir)]})
						else:
							self.recipients[recipient].append(self.format_filename(ed_dir))
						pass
						self.write_member(a_id, row, ed_fid, recipient)
					pass
				pass
			pass
			#
			# update the distribution lists--only regional members will be in the distribution list
			# and only for the list marked validate="True" in the XML
			#
#			if ((reg_fid == this_reg_fid) and not region_only and self.validate):
			if ((reg_fid == this_reg_fid) and self.validate):
#				if (a_id == "L02711 10"):
#					print("Found: %s\n" % (a_id))
				if (not distribution.is_member_validated(a_id)):
#					if (a_id == "L02711 10"):
#						print("Validating: %s\n" % (a_id))
					distribution.validate_recipient(self.nmra_map, parent_dir, dist_dir, zip_filename, a_id, r_region, r_division, r_lname, r_fname, a_email, force_override)
				pass
			pass
		pass
		#
		# at the end of the input sheet, write out all of the output sheets we made
		#
		self.save_workbooks()
		return(self.recipients)
	pass
pass

