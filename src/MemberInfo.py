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
	def __init__(self, legacy_mode):
		self.legacy_mode = legacy_mode
		self.column_headers = []
		self.member = {}
		self.alt_required_columns = ['id', 'memberregion', 'division', 'lname', 'fname']
		self.leg_alt_required_columns = ['id', 'memberregion', 'division', 'lname', 'fname', 'email']
		self.reg_required_columns = ['id', 'region', 'division', 'lname', 'fname']
		self.leg_reg_required_columns = ['id', 'region', 'division', 'lname', 'fname', 'email']
		if (legacy_mode):
			self.required_columns = self.leg_reg_required_columns
		else:
			self.required_columns = self.reg_required_columns
		pass
	pass
	#
	# this is the header information in each report file as it is processed
	#
	def add_column_header(self, name):
		self.column_headers.append(name)
	pass
	#
	# this checks that the header contains the required columns needed for processing
	#
	def has_valid_header(self):
		ret_val = True
		if ('memberregion' in self.column_headers):
			if (self.legacy_mode):
				self.required_columns = self.leg_alt_required_columns
			else:
				self.required_columns = self.alt_required_columns
			pass
		else:
			if (self.legacy_mode):
				self.required_columns = self.leg_reg_required_columns
			else:
				self.required_columns = self.reg_required_columns
			pass
		pass
		for name in self.column_headers:
			if not (name.lower() in self.required_columns):
				ret_val = False
			pass
		pass
		return ret_val
	pass
	#
	# these methods return the different pieces of the header information
	#
	def get_column(self, df, row, name):
		try:
			retval = df.at[row, name]
		except ValueError:
			raise("Can't find column '%s'" % (name))
		pass
		return retval
	pass
#	def get_name(self, position):
#		return self.column_array[position]
#	pass
	#
	# return true if the given column exists
	#
	def has_column(self, name):
		if (name.lower() in [x.lower() for x in self.column_headers]):
			return True
		else:
			return False
		pass
	pass
	#
	# this is a place to temporarily store required member information during processing
	#
	def set_member(self, df, row):
		a_id		= '%s' % self.get_column(df, row, 'id')
		a_lname		= '%s' % self.get_column(df, row, 'lname')
		a_fname		= '%s' % self.get_column(df, row, 'fname')
		t_division	=    self.get_column(df, row, 'division')
		try:
			a_division	=    int(t_division)
		except ValueError:
			a_division  = 0
		pass
		t_region	=    self.get_column(df, row, 'region')
		try:
			a_region	=    int(t_region)
		except ValueError:
			a_region    = 0
		pass

		if (self.has_column('email')):
			a_email		= '%s' % self.get_column(df, row, 'email')
		else:
			a_email		= ''
		pass
		self.member = {'id' : a_id, 'region' : a_region, 'division' : a_division, 'lname' : a_lname, 'fname' : a_fname, 'email' : a_email}
		print("Member Info:")
		for key in self.member.keys():
			print("\t%s =>%s" % (key, self.member[key]))
		pass
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

