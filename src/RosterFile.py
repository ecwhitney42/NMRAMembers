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
import pyexcel
import pyexcel_xlsx
import MemberInfo
import EmailDistribution
import xlrd
import xlwt
import os
import re

class RosterFile:
	#
	# Constructor
	#
	def __init__(self, work_dir, nmra_map, use_long, reassignments):
		self.work_dir = work_dir
		self.nmra_map = nmra_map
		self.use_long = use_long
		self.reassignments = reassignments
		self.member_info = MemberInfo.MemberInfo()
		self.sheet_name = ""
		self.region_filenames = {}
		self.regions = []
		self.division_filenames = {}
		self.divisions = []
		self.report_name = ""
		self.roster_wb = {}
		self.roster_ws = {}
		self.roster_wb_row = {}
		self.roster_rb = None
		self.roster_rs = None
		self.roster_nrows = 0
		self.roster_ncols = 0
	pass
	#
	# this method returns true if the given member ID is in the list of reassignments 
	# and the first and last name matches and the from division is not the to division 
	#
	def is_member_reassigned(self, nmra_id):
		ret_val = False
		if (self.reassignments.has_member(nmra_id)):
			if ((self.member_info.get_division() == self.reassignments.get_from_division(nmra_id)) and (self.member_info.get_division() != self.reassignments.get_to_division(nmra_id))):
				if ((self.member_info.get_lname() == self.reassignments.get_lname(nmra_id)) and (self.member_info.get_fname() == self.reassignments.get_fname(nmra_id))):
					ret_val = True
				else:
					print("WARNING: Reassignment found for member %s, however, the name doesn't match so no reassignment" % nmra_id)
				pass
			pass
		pass
		return ret_val
	pass
	#
	# Read the given roster file and save the information necessary to create a new workbook from it.
	# This is where we are able to read in the crusty old version of Excel files.
	#
	def read_file(self, file_name):
		#
		# Read in the .xls in its crusty old Excel format--this may issue warnings
		# but they can be ignored
		#
		try:
			self.roster_rb = xlrd.open_workbook(filename=file_name, encoding_override="cp1252", formatting_info = True)
		except:
			print("Roster File Error: ", sys.exc_info()[0])
			raise
		pass
		#
		# Get the sheet
		#
		self.roster_rs = self.roster_rb.sheet_by_index(0)

		self.roster_ncols = self.roster_rs.ncols
		self.roster_nrows = self.roster_rs.nrows
		#
		# take out the extra numbers in the filename and save the root string
		# to a report_name that way we can write out cleaner filenames
		#
		report_file = os.path.splitext(os.path.basename(file_name))
		self.report_name = report_file[0]
		re1 = re.compile(r'^(\d+)_(.*)')
		self.report_name = re1.sub(r'\2', self.report_name)
		re2 = re.compile(r'(.*)(\d+)$')
		self.report_name = re2.sub(r'\1', self.report_name)
		#
		# Look at the header to get the MemberInfo
		#
		for col in range(0, self.roster_ncols):
			self.member_info.add_column_header(self.roster_rs.cell(0, col).value, col)
		pass
		if (not self.member_info.has_valid_header()):
			raise ValueError("%s doesn't have the required headers" % (file_name))
		pass
		#
		# This fixes a problem with long sheet names that the NMRA uses with their
		# crusty old format--just truncate the sheet name to 31 characters
		#
		for sn in self.roster_rb.sheet_names():
			if (len(sn) > 31):
				self.sheet_name = sn[0:30]
			else:
				self.sheet_name = sn
			pass
		pass
	pass
	#
	# Create a workbook for a given output file.
	# This uses the xlwt library to create the workbook
	#
	def create_workbook(self, nmra_id, dir_path):
		os.makedirs("%s" % dir_path, exist_ok=True)
		self.division_filenames[nmra_id] = "%s/%s.xls" % (dir_path, self.report_name)
		#
		# Create a new workbook to copy it into for processing
		#
		self.roster_wb[nmra_id] = xlwt.Workbook(style_compression=2)
		self.roster_ws[nmra_id] = self.roster_wb[nmra_id].add_sheet(self.sheet_name, cell_overwrite_ok=True)
		#
		# insert the header row into this worksheet
		#
		if (not nmra_id in self.roster_wb_row.keys()):
			self.roster_wb_row[nmra_id] = 1
			for col in range(0, self.roster_ncols):
				self.roster_ws[nmra_id].write(0, col, self.member_info.get_name(col))
			pass	
		pass
	pass
	#
	# Copy member info from current worksheet to given output worksheet
	# This has to look for dates and zip codes and handle their formatting explicitly
	#
	def write_row(self, file_id, row):
		for col in range(0, self.roster_ncols):
			cell_value = self.roster_rs.cell_value(row, col)
			cell_type = self.roster_rs.cell_type(row, col)
			if (cell_type == 3):
				self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, cell_value, xlwt.Style.easyxf(num_format_str="mm/dd/yyyy"))
			elif (cell_type == 2 and self.member_info.get_name(col)=="zip"):
				self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, cell_value, xlwt.Style.easyxf(num_format_str="00000"))
			else:
				self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, cell_value)
			pass
		pass
		self.roster_wb_row[file_id] = self.roster_wb_row[file_id] + 1
	pass
	#
	# Copy member info from current worksheet to given output worksheet
	# and reassign this member to the given region/division if necessary
	#
	def reassign_member(self, file_id, row, new_region, new_division):
		r_col = self.member_info.get_column('region')
		d_col = self.member_info.get_column('division')
		for col in range(0, self.roster_ncols):
			cell_value = self.roster_rs.cell_value(row, col)
			cell_type = self.roster_rs.cell_type(row, col)
			if (col == r_col):
				self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, "%02d" % new_region)
			elif (col == d_col):
				self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, "%02d" % new_division)
			else:
				if (cell_type == 3):
					self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, cell_value, xlwt.Style.easyxf(num_format_str="mm/dd/yyyy"))
				elif (cell_type == 2 and self.member_info.get_name(col)=="zip"):
					self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, cell_value, xlwt.Style.easyxf(num_format_str="00000"))
				else:
					self.roster_ws[file_id].write(self.roster_wb_row[file_id], col, cell_value)
				pass
			pass
		pass
		self.roster_wb_row[file_id] = self.roster_wb_row[file_id] + 1
	pass
	#
	# Write a member to the given row and check if they are reassigned
	#
	def write_member(self, nmra_id, row, division, region_id, division_id):
		m_region = self.member_info.get_region()
		m_division = self.member_info.get_division()
		#
		# check to see if this member is reassigned or not
		#
		reassign = self.is_member_reassigned(nmra_id)
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
		if (division > 0):
			if (reassign):
				print("\tProcessing division reassignment for NMRA member %s from division %02d%02d to division %02d%02d" % (nmra_id, m_region, m_division, r_region, r_division))
				self.reassign_member(division_id, row, r_region, r_division)
			else:
				self.write_row(division_id, row)
			pass
		pass
		#
		# Write all members to their appropriate division
		#
		if (reassign):
			print("\tProcessing division reassignment for NMRA member %s from division %02d%02d to division %02d%02d" % (nmra_id, m_region, m_division, r_region, r_division))
			self.reassign_member(region_id, row, r_region, r_division)
		else:
			self.write_row(region_id, row)
		pass
	pass

	#
	# Save the current output workbooks
	#
	def save_workbooks(self):
		for div_fid in self.division_filenames.keys():
			self.roster_wb[div_fid].save(self.division_filenames[div_fid])
		pass
		for reg_fid in self.region_filenames.keys():
			self.roster_wb[reg_fid].save(self.region_filenames[reg_fid])
		pass
	pass
	#
	# Processes the current workbook in memory
	#
	def process(self, distribution):
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
		for row in range(1, self.roster_nrows):
			#
			# The first row contains the column headings we need to find the offsets
			#
			ncol = self.roster_ncols
			self.member_info.set_member(self.roster_rs, row)

			a_id		= self.member_info.get_id()
			a_region	= self.member_info.get_region()
			a_division	= self.member_info.get_division()
			a_lname		= self.member_info.get_lname()
			a_fname		= self.member_info.get_fname()
			a_email		= self.member_info.get_email()
			#
			# Create a workbook for each division and region encountered
			#
			# If it's a just a region entry, put it in a region file otherwise break out by division
			#
			# This converts the 4-digit _fid code to a text _name string if it's in the division map
			#
			reg_fid = self.nmra_map.get_file_id(a_region, 0)
			div_fid = self.nmra_map.get_file_id(a_region, a_division)
			if (self.reassignments.has_member(a_id)):
				r_region	= self.reassignments.get_to_region(a_id)
				r_division	= self.reassignments.get_to_division(a_id)
				r_lname		= self.reassignments.get_lname(a_id)
				r_fname		= self.reassignments.get_fname(a_id)
				if (self.is_member_reassigned(a_id)):
					reg_fid = self.nmra_map.get_file_id(r_region, 0)
					div_fid = self.nmra_map.get_file_id(r_region, r_division)
				pass
			else:
				r_division	= a_division
				r_region	= a_region
				r_lname		= a_lname
				r_fname		= a_fname
			pass
			#
			# Make sure we have mapped this region/division
			#
			if self.nmra_map.has_region_id(reg_fid):
				#
				# The difference between long and short region names comes from
				# the NMRA Region/Division MAP file that has the RID column
				#
				if self.use_long:
					reg_name = self.nmra_map.get_region(reg_fid)
				else:
					reg_name = self.nmra_map.get_region_id(reg_fid)
				pass
			else:
				raise ValueError("Region ID for %s not found in the map" % reg_fid)
			pass
			if self.nmra_map.has_division(div_fid):
				div_name = self.nmra_map.get_division(div_fid)
			else:
				raise ValueError("Division for %s not found in the map", div_fid)
			pass
			#
			# Assign the output file to the correct region or division
			#
			if ((not div_fid in self.divisions) and (r_division > 0)):
				self.divisions.append(div_fid)
				if self.use_long:
					div_dir = "%s/%s_Region-%s_Division" % (self.work_dir, reg_name, div_name)
				else:
					div_dir = "%s/%s_Region-%s_Division" % (self.work_dir, reg_name, div_name)
				pass
				self.create_workbook(div_fid, div_dir)
			pass
			if (not reg_fid in self.regions):
				self.regions.append(reg_fid)
				if self.use_long:
					reg_dir = "%s/%s_Region" % (self.work_dir, reg_name)
				else:
					reg_dir = "%s/%s_Region" % (self.work_dir, reg_name)
				pass
				self.create_workbook(reg_fid, reg_dir)
			#
			# update the division and region entries
			#
			self.write_member(a_id, row, r_division, reg_fid, div_fid)
			distribution.validate_recipient(a_id, r_region, r_division, r_lname, r_fname, a_email)
		pass
		#
		# at the end of the input sheet, write out all of the output sheets we made
		#
		self.save_workbooks()
	pass
pass

