#------------------------------------------------------------------------------
#
# Class MemberInfo
#
#------------------------------------------------------------------------------
#
# This is a class created to contain some useful information about the member
# data in the roster. 
#
#------------------------------------------------------------------------------
class MemberInfo:
	#
	# Constructor
	#
	def __init__(self):
		self.column_headers = {}
		self.column_array = []
		self.member = {}
		self.required_columns = {'id' : -1, 'lname' : -1, 'fname' : -1, 'region' : -1, 'division' : -1}
	pass
	#
	# this is the header information in each report file as it is processed
	#
	def add_column_header(self, name, position):
		self.column_headers[name] = position
		self.column_array.append(name)
		if (name in self.required_columns.keys()):
			self.required_columns[name] = position
		pass
	pass
	#
	# this checks that the header contains the required columns needed for processing
	#
	def has_valid_header(self):
		ret_val = True
		for name in self.required_columns.keys():
			if (self.required_columns == -1):
				ret_val = False
			pass
		pass
		return ret_val
	pass
	#
	# these methods return the different pieces of the header information
	#
	def get_column(self, name):
		return self.column_headers[name]
	pass
	def get_name(self, position):
		return self.column_array[position]
	pass
	#
	# return true if the given column exists
	#
	def has_column(self, name):
		if (name in self.column_headers.keys()):
			return True
		else:
			return False
		pass
	pass
	#
	# this is a place to temporarily store required member information during processing
	#
	def set_member(self, sheet, row):
		a_id		= '%s' % sheet.cell(row, self.get_column('id')).value
		a_lname		= '%s' % sheet.cell(row, self.get_column('lname')).value
		a_fname		= '%s' % sheet.cell(row, self.get_column('fname')).value
		t_division	=    sheet.cell(row, self.get_column('division')).value
		t_region	=    sheet.cell(row, self.get_column('region')).value
		if (self.has_column('email')):	
			a_email		= '%s' % sheet.cell(row, self.get_column('email')).value
		else:
			a_email		= ""
		pass
		try:
			a_division	=    int(t_division)
		except ValueError:
			a_division  = 0
		try:
			a_region	=    int(t_region)
		except ValueError:
			a_region    = 0
			
		self.member = {'id' : a_id, 'region' : a_region, 'division' : a_division, 'lname' : a_lname, 'fname' : a_fname, 'email' : a_email}
	pass
	#
	# these methods return the difference pieces of the member information
	#
	def get_member(self, nmra_id):
		return self.member
	pass
	def get_id(self):
		return self.member['id']
	pass
	def get_region(self):
		return self.member['region']
	pass
	def get_division(self):
		return self.member['division']
	pass
	def get_lname(self):
		return self.member['lname']
	pass
	def get_fname(self):
		return self.member['fname']
	pass
	def get_email(self):
		return self.member['email']
	pass
pass

