#------------------------------------------------------------------------------
#
# Class MemberReassignments
#
# The file format for the reassignment file is a single sheet in .xlsx
# format with the following required columns:
#
# id, region, division, lname, fname
#
# id:				the NMRA number of the member to reassign
# from_region:		the two-digit NMRA region number to reassign this member from
# from_division:	the two-digit NMRA division number to reassign this member from
# to_region:		the two-digit NMRA region number to reassign this member to
# to_division:		the two-digit NMRA division number to reassign this member to
# lname:			the last name of the member
# fname:			the first name of the member
#------------------------------------------------------------------------------
#
# This class manages the member reassignment information
#
#------------------------------------------------------------------------------
import sys
import pyexcel
import pyexcel_xlsx

class MemberReassignments:
	#
	# Constructor
	#
	def __init__(self):
		self.members = {}
		self.required_columns = {'id' : -1, 'from_region' : -1, 'from_division' : -1, 'to_region' : -1, 'to_division' : -1, 'lname' : -1, 'fname' : -1}
	pass
	#
	# Return the given field from the reassignment list for the given NMRA member
	# these assume has_member() has been called to check for this NMRA member first
	#	
	def get_from_region(self, nmra_id):
		return self.members[nmra_id]['from_region']
	pass
	def get_from_division(self, nmra_id):
		return self.members[nmra_id]['from_division']
	pass
	def get_to_region(self, nmra_id):
		return self.members[nmra_id]['to_region']
	pass
	def get_to_division(self, nmra_id):
		return self.members[nmra_id]['to_division']
	pass
	def get_lname(self, nmra_id):
		return self.members[nmra_id]['lname']
	pass
	def get_fname(self, nmra_id):
		return self.members[nmra_id]['fname']
	pass
	def get_member(self, nmra_id):
		return self.members[nmra_id]
	pass
	#
	# Return true of the given member is in the reassignment list
	#
	def has_member(self, nmra_id):
		if (nmra_id in self.members.keys()):
			return True
		else:
			return False
		pass
	pass
	#
	# get the column number for the given column
	#
	def get_column(self, name):
		return self.required_columns[name]
	pass
	#
	# Add a member to the reassignment list
	#
	def add_member(self, nmra_id, from_region, from_division, to_region, to_division, lname, fname):
		self.members[nmra_id] = { 'from_region' : from_region, 'from_division' : from_division, 'to_region' : to_region, 'to_division' : to_division, 'lname' : lname, 'fname' : fname}
	pass
	#
	# Reads the reassignment file and populates its data structure in memory
	#
	def read_file(self, filename):
		print("Reading the NMRA Member Division Reassignment File: %s" % filename)
		try:
			reassign_ws = pyexcel.get_sheet(file_name=filename)
		except:
			print("Division Reassignment Error: ", sys.exc_info()[0])
			raise
		pass

		all_good = True
		row_num = 0
		for row in reassign_ws:
			#
			# The first row contains the column headings we need to find the offsets
			#
			if (row_num == 0):
				col_num = 0
				for cell in row:
					for key in self.required_columns.keys():
						if (cell == key):
							self.required_columns[key] = col_num
						pass
					col_num = col_num + 1
				pass
				for key in self.required_columns.keys():
					if (self.required_columns[key] == -1):
						all_good = False
					pass
				pass
				if (all_good == False):
					raise ValueError('All required columns MUST be included in the Division Reassignment file!')
				pass
			else:
				r_id       = "%s" % row[self.required_columns['id']]
				f_region   = int(row[self.required_columns['from_region']])
				f_division = int(row[self.required_columns['from_division']])
				t_region   = int(row[self.required_columns['to_region']])
				t_division = int(row[self.required_columns['to_division']])
				r_lname    = "%s" % row[self.required_columns['lname']]
				r_fname    = "%s" % row[self.required_columns['fname']]
				if (f_region != t_region):
					raise ValueError('NMRA Members are NOT allowed to change regions with this program!')
				pass
				print("Reassigning NMRA Member %s, (%s, %s) from Division %02d%02d to Division %02d%02d" % (r_id, r_lname, r_fname, f_region, f_division, t_region, t_division))

				self.add_member(r_id, f_region, f_division, t_region, t_division, r_lname, r_fname)
			pass
			row_num = row_num + 1	
		pass
	pass
pass

