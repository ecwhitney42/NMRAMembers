
#!/usr/bin/env python3
import xml.etree.ElementTree as ET
import sys
#
# Class: NMRAMembersConifig
#
# This class is for reading and managing the XML file used to configure
# the datasets that come from NMRA HQ for the NMRAMembers program.
#
class NMRAMembersConfig:
	#
	# Element names from the XML file
	#
	ELEM_DEFS			= 'definitions'
	ELEM_PARAMS			= 'parameter'
	ELEM_NAME			= 'name'
	ELEM_VALIDATE		= 'validate'
	ELEM_VALUE			= 'value'
	ELEM_DESCRIPTION	= 'description'
	ELEM_RPTS			= 'reports'
	ELEM_FILE			= 'file'
	ELEM_ACTION			= 'action'
	ELEM_RECIP			= 'recipient'
	ELEM_ROSTER_FORMAT	= 'format'
	ELEM_ROSTER_FIELDS	= 'fields'
	ELEM_IFORMAT		= 'iformat'
	ELEM_OFORMAT		= 'oformat'
	ELEM_OUTPUT			= 'output'
	ELEM_INCLUDE		= 'include'
	ELEM_SORT			= 'sort'
	ELEM_DATE_FORMAT	= 'date_format'
	ELEM_DATE_FIELDS	= 'date_fields'
	
	ACT_NEWSLETTER		= 'newsletter'
	ACT_COPY			= 'copy'
	ACT_REASSIGN		= 'reassignment'
	#
	# Constructor takes the XML file name and which mode to use from that file
	#
	def __init__(self, filename, mode, verbose=False):
		self.xml_tree = None
		self.xml_config = None
		self.xml_mode = mode
		self.action_list = []
		self.valid_recipients = []
		self.valid_actions = []
		self.output_format=""
		self.input_format=""
		self.roster_format=""
		self.date_format=""
		self.reports = {}
		self.instance_list = {}
		self.validate = {}
		
		if (verbose):
			print("Reading the XML Configuration File %s..." % (filename))
		pass
		try:
			self.xml_tree = ET.parse(filename)
			self.xml_config = self.xml_tree.getroot()
		except:
			print("XML Configuration File Error: ", sys.exc_info()[0])
			raise
		pass
		if (verbose):
			print("\tValid XML parameters in %s" % (filename))
		pass
		for parameter in self.xml_config.findall('%s/%s/%s' % (self.xml_mode, self.ELEM_DEFS, self.ELEM_PARAMS)):
			name  = parameter.get(self.ELEM_NAME)
			value = parameter.get(self.ELEM_VALUE)
			desc  = parameter.get(self.ELEM_DESCRIPTION)		
			if (verbose):
				print("\t\t%9s = %12s: %s" % (name, value, desc))
			pass
			if (name == self.ELEM_RECIP):
				self.valid_recipients.append(value)
			elif (name == self.ELEM_ACTION):
				self.valid_actions.append(value)
			elif (name == self.ELEM_OFORMAT):
				self.output_format = value
			elif (name == self.ELEM_IFORMAT):
				self.input_format = value
			elif (name == self.ELEM_DATE_FORMAT):
				self.date_format = value
			else:
				print("Unknown parameter...")
			pass
		pass
		if (verbose):
			print("\n\n")
		pass

		for report in self.xml_config.findall('%s/%s/*' % (self.xml_mode, self.ELEM_RPTS)):
			action = report.tag
			self.action_list.append(action)
			if (action == self.ACT_NEWSLETTER):
				for file in self.xml_config.findall('%s/%s/%s/%s' % (self.xml_mode, self.ELEM_RPTS, action, self.ELEM_FILE)):
					filename = file.get('name')
#					print("Configuring file: %s" % (filename))
					if filename in self.instance_list.keys():
						self.instance_list[filename] = self.instance_list[filename] + 1
					else:
						self.instance_list[filename] = 1
					pass
					recip_list = []
					param_list= {}
					roster_field_list = ""		# NOTE: the mapping between field and format is current implied. Code to make this
					roster_format_list = ""	# configurable will need to be added and the XML will need to be updated.
					include_list = ""
					sort_list = {}
					output_list = {}
					output_name = filename
					for param in file.findall('*'):
						if (param.tag == self.ELEM_RECIP):
							recip_list.append(param.text)
						elif (param.tag == self.ELEM_ROSTER_FIELDS):
							roster_field_list = param.text.strip('"')
						elif (param.tag == self.ELEM_ROSTER_FORMAT):
							roster_format_list = param.text.strip('"')
						elif (param.tag == self.ELEM_OUTPUT):
							output_name = param.get('name')
#							print("output_name = %s" % (output_name))
							for incl in file.findall('*/%s' % (self.ELEM_INCLUDE)):
								include_list = incl.text.strip('"')
#								print("Include: %s" % (include_list))
							pass
						elif (param.tag == self.ELEM_SORT):
							sort_list = {self.ELEM_SORT : param.text}
						pass
					pass
					param_list.update({'instance' 				: self.instance_list[filename]})
					param_list.update({self.ELEM_RECIP   		: recip_list})
					param_list.update({self.ELEM_ROSTER_FIELDS  : roster_field_list})
					param_list.update({self.ELEM_ROSTER_FORMAT 	: roster_format_list})
					param_list.update({self.ELEM_OUTPUT  		: output_name})
					param_list.update({self.ELEM_INCLUDE 		: include_list})
					param_list.update({self.ELEM_SORT    		: sort_list})
					param_key = "%s#%s" % (filename, self.instance_list[filename])
#					print("Updating param_key: %s for file: %s" % (param_key, file))
					params = {param_key : param_list}
					if action in self.reports.keys():
						self.reports[action].update(params)
					else:
						self.reports[action] = params
					pass
				pass
			elif ((action == self.ACT_COPY) or (action == self.ACT_REASSIGN)):
				for file in self.xml_config.findall('%s/%s/%s/%s' % (self.xml_mode, self.ELEM_RPTS, action, self.ELEM_FILE)):
					filename = file.get('name')
					validate = file.get('validate')
					if (type(validate) != 'NoneType'):
						self.validate.update({filename : validate})
					pass
					if filename in self.instance_list.keys():
						self.instance_list[filename] = self.instance_list[filename] + 1
					else:
						self.instance_list[filename] = 1
					pass
					param_list= {}
					recip_list = []
					date_field_list = []
					for param in file.findall('*'):
						if (param.tag == self.ELEM_RECIP):
							recip_list.append(param.text)
						elif (param.tag == self.ELEM_DATE_FIELDS):
							date_field_string = param.text.strip('"')
							date_field_list = date_field_string.split(',')
						pass
					pass
#					print("Action: %s" % (action))
#					print("Filename: %s" % (filename))
#					print("Adding: 'instance' : %d" % self.instance_list[filename])
#					print("\t%s : %s" % (self.ELEM_RECIP, recip_list))
#					print("\t%s : %s" % (self.ELEM_DATE_FIELDS, date_field_list))
					param_list.update({'instance' : self.instance_list[filename]})
					param_list.update({self.ELEM_RECIP : recip_list})
					param_list.update({self.ELEM_DATE_FIELDS : date_field_list})
					param_key = "%s#%s" % (filename, self.instance_list[filename])
					params = {param_key : param_list}
					if action in self.reports.keys():
						self.reports[action].update(params)
					else:
						self.reports[action] = params
					pass
				pass
			else:
				print("Unrecognized action '%s' found in XML..." % (action))
			pass
		pass
	
		if (not self.validate_recipients()):
			print("XML contains invalid recipients!")
		pass
	
		if (not self.validate_actions()):
			print("XML contains invalid actions!")
		pass
	pass
	#
	# Returns True if the given file has a validate flag set
	#
	def get_validate(self, filename):
		if (filename in self.validate.keys()):
			return self.validate.get(filename)
		else:
			return False
		pass
	pass
	#
	# Returns the list of actions from the XML
	#
	def get_action_list(self):
		return(self.action_list)
	pass
	#
	# Returns the input format parameter from the XML
	#
	def get_input_format(self):
		return(self.input_format)
	pass
	#
	# Returns the output format parameter from the XML
	#
	def get_output_format(self):
		return(self.output_format)
	pass
	#
	# Returns the date format parameter from the XML
	#
	def get_roster_format(self):
		return(self.roster_format)
	pass
	#
	# Returns the date format parameter from the XML
	#
	def get_date_format(self):
		return(self.date_format)
	pass
	#
	# Returns the list of files for the given action from the XML
	#
	def get_files(self, action):
		file_list = []
		for file in self.xml_config.findall("%s/%s/%s/%s" % (self.xml_mode, self.ELEM_RPTS, action, self.ELEM_FILE)):
			filename = file.get('name')
			file_list.append("%s" % (filename))
		pass
		return(file_list)
	pass
	#
	# Returns a list of parameters for the given report under the action and filename
	#
	def get_report_parameter_names(self, action, filename, instance):
		params = []
		for param in self.reports[action][filename][instance].keys():
			for (param_tag, param_value) in self.reports[action][filename][instance]:
				param.append(param_tag)
			pass
		pass
		return(params)
	pass
	#
	# Returns the value of the parameters for the given report under the action, filename and parameter name
	#
	def get_report_parameter_value(self, action, filename, instance, name):
		value = ""
		insthash = "%s#%s" % (filename, instance)
		if insthash in self.reports[action].keys():
			param = self.reports[action][insthash]
			if name in param.keys():
				value = param[name]
			pass
		pass
		return(value)
	pass
	#
	# Returns the list of instances for the given filename
	#
	def get_report_instance_list(self, action, filename):
		instances = []
		for params in self.reports[action].keys():
			(name, inst) = params.split('#')
			if (name == filename):
				instances.append(int(inst))
			pass
		pass
		return(instances)
	pass
	#
	# Validate the recipients in the XML against the ones defined in the parameter section of the XML
	#
	def validate_recipients(self):
		all_good=True
		for file in self.xml_config.findall("%s/%s/*/%s" % (self.xml_mode, self.ELEM_RPTS, self.ELEM_FILE)):
			for recipient in file.findall('./%s' % (self.ELEM_RECIP)):
				if (not recipient.text in self.valid_recipients):
					print("%s not valid" % (recipient.text))
					all_good = False
				pass
			pass
		pass
		return(all_good)
	pass
	#
	# Validate the actions in the XML against the ones defined in the parameter section of the XML
	#
	def validate_actions(self):
		all_good=True
		for file in self.xml_config.findall("%s/%s/*/*" % (self.xml_mode, self.ELEM_RPTS)):
			for action in file.findall('./%s' % (self.ELEM_ACTION)):
				if (not action in self.valid_actions):
					all_good = False
				pass
			pass
		pass
		return(all_good)
	pass
	#
	# Returns the list of recipients for the given file for the action provided
	#
	def get_recipients(self, action, filename, instance):
		recipient_list = []
		insthash = "%s#%s" % (filename, instance)
		if insthash in self.reports[action].keys():
			param = self.reports[action][insthash]
			if (self.ELEM_RECIP in param.keys()):
#				print("PARAM[self.ELEM_RECIP] IS: %s" % (param.get(self.ELEM_RECIP)))
				recipient_list = param.get(self.ELEM_RECIP)
			pass
		pass
#		print("Recipients for '%s' in '%s(%d)' are: %s" % (action, filename, instance, recipient_list))
		return(recipient_list)
	pass
	#
	# Returns the list of includes for the given file for the action provided
	#
	def get_includes(self, action, filename, instance):
		include_list = []
		insthash = "%s#%s" % (filename, instance)
		if insthash in self.reports[action].keys():
			param = self.reports[action][insthash]
			include_list = param.get(self.ELEM_INCLUDE)
		pass
#		print("Includes are: %s" % (include_list))
		return(include_list)
	pass
	#
	# Returns the list of sorts for the given file for the action provided
	#
	def get_sorts(self, action, filename, instance):
		sort_list = {}
		insthash = "%s#%s" % (filename, instance)
		if insthash in self.reports[action].keys():
			param = self.reports[action][insthash]
			sort_list = param.get(self.ELEM_SORT)
		pass
#		print("Sorts are: %s" % (sort_list))
		return(sort_list)
	pass
	#
	# Returns the list of recipients for the given file for the action provided
	#
	def get_outputs(self, action, filename, instance):
		output_list = []
		insthash = "%s#%s" % (filename, instance)
		if insthash in self.reports[action].keys():
			param = self.reports[action][insthash]
			if (self.ELEM_OUTPUT in param.keys()):
				output_list.append(param.get(self.ELEM_OUTPUT))
			pass
		pass
#		print("Outputs are: %s" % (output_list))
		return(output_list)
	pass
	#
	# Returns the list of date fields for the given file for the action provided
	#
	def get_date_fields(self, action, filename, instance):
		date_field_list = []
		insthash = "%s#%s" % (filename, instance)
		if insthash in self.reports[action].keys():
			param = self.reports[action][insthash]
#			print("Param in get_date_fields() is %s" % (param))
			if (self.ELEM_DATE_FIELDS in param.keys()):
				date_field_list = param.get(self.ELEM_DATE_FIELDS)
			pass
		pass									   
#		print("Date Fields for '%s' in '%s(%d)' are: %s" % (action, filename, instance, date_field_list))
		return(date_field_list)
	pass
pass

