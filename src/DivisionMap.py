###############################################################################
#------------------------------------------------------------------------------
#
# Class DivisionMap
#
#------------------------------------------------------------------------------
# The file format for the map file is a single sheet in .xlsx
# format with the following required columns:
#
# region, division, name, RID
#
# region:	the two-digit numerical value of the region as defined on the NMRA 
# 			web site
# division:	the two-digit numerical value of the division as defined on the 
# 			NMRA web site
# name:		the name of the division as defined on the NMRA web site
#			the region name is given in this field if the division is set to 0
# RID:		the region nickname, often give as 3 upper case letters as defined
# 			on the NMRA web site
#------------------------------------------------------------------------------
#
# This class manages the mapping of the 4-digit region/division ID to text 
# strings based on the information found on the NMRA web site.
#
#------------------------------------------------------------------------------
import sys
import pandas as pd

class DivisionMap:
	#
	# Constructor
	#
	def __init__(self):
		self.region_map = {}
		self.region_ids = {}
		self.division_map = {}
		self.required_columns = ['region', 'division', 'name', 'RID']

	pass
	#
	# Returns file IDs used for making hashes
	#
	def get_file_id(self, region, division):
		return "%02d%02d" % (region, division)
	pass
	#
	# Returns the region name for the given region number
	# 	
	def get_region(self, file_id):
		return self.region_map[file_id]
	pass
	#
	# Returns the region ID (nickname) for the given region number
	#
	def get_region_id(self, file_id):
		return self.region_ids[file_id]
	pass
	#
	# Returns the divsion name for the given division number
	#
	def get_division(self, file_id):
		return self.division_map[file_id]
	pass
	#
	# Returns true if the given region number has a region ID (nickname)
	#  
	def has_region_id(self, file_id):
		if (file_id in self.region_ids.keys()):
			return True
		else:
			return False
		pass
	pass
	#
	# Returns true if the given region number has a region name defined
	#
	def has_region(self, file_id):
		if (file_id in self.region_map.keys()):
			return True
		else:
			return False
		pass
	pass
	#
	# Returns true if the given division number has a division name defined
	# 	
	def has_division(self, file_id):
		if (file_id in self.division_map.keys()):
			return True
		else:
			return False
		pass
	pass
	#
	# Reads the map file and populates its data structure in memory with the contents
	#
	def read_file(self, filename):
		print("Reading the NMRA Region/Division Map File: %s" % filename)
		try:
			excel_file = pd.ExcelFile(filename)
		except:
			print("Division Map Error: ", sys.exc_info()[0])
			raise
		pass
		#
		# read in the division map
		#
		data_frame = excel_file.parse()
		header = data_frame.columns
		#
		# first, check to make sure that all of the required columns are present in the file
		#
		all_good = True
		for name in self.required_columns:
			if not (name.lower() in [x.lower() for x in header]):
				all_good = False
			pass
		pass
		if (all_good == False):
			raise ValueError('All required columns MUST be included in the Division Map file!')
		pass
		for row, data in data_frame.iterrows():
			dm_region   = int(data['region'])
			dm_division = int(data['division'])
			dm_name     = "%s" % data['name']
			dm_name		= dm_name.replace(" ", "_")
			dm_rid		= "%s" % data['RID']
			div_fid		= self.get_file_id(dm_region, dm_division)
			reg_fid		= self.get_file_id(dm_region, 0)
#			print("Mapping ID %02d%02d to %s Region %s Division" % (dm_region, dm_division, dm_rid, dm_name))
			if (dm_division == 0):
				self.region_map[reg_fid] = dm_name
				self.region_ids[reg_fid] = dm_rid
			pass
			self.division_map[div_fid] = dm_name
		pass
	pass
pass

