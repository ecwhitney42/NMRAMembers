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
import pandas as pd

class MemberReassignments:
	#
	# Constructor
	#
	def __init__(self):
		self.members = {}
		self.required_columns = ['id', 'from_region', 'from_division', 'to_region', 'to_division', 'lname', 'fname']
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
			excel_file = pd.ExcelFile(filename)
		except:
			print("Division Reassignment Error: ", sys.exc_info()[0])
			raise
		pass

		data_frame = excel_file.parse()
		header = data_frame.columns

		all_good = True
		for name in self.required_columns:
			if not (name.lower() in [x.lower() for x in header]):
				all_good = False
			pass
		pass	
		if (all_good == False):
			raise ValueError('All required columns MUST be included in the Division Reassignment file!')
		pass
		for row, data in data_frame.iterrows():
			r_id       = "%s" % data['id']
			f_region   = int(data['from_region'])
			f_division = int(data['from_division'])
			t_region   = int(data['to_region'])
			t_division = int(data['to_division'])
			r_lname    = "%s" % data['lname']
			r_fname    = "%s" % data['fname']
			if (f_region != t_region):
				raise ValueError('NMRA Members are NOT allowed to change regions with this program!')
			pass
			print("Reassigning NMRA Member %s, (%s, %s) from Division %02d%02d to Division %02d%02d" % (r_id, r_lname, r_fname, f_region, f_division, t_region, t_division))

			self.add_member(r_id, f_region, f_division, t_region, t_division, r_lname, r_fname)
		pass
	pass
pass

