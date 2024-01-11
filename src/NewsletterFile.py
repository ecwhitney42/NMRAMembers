#------------------------------------------------------------------------------
#
# Class NewsletterFile
#
# Process a newsletter roster file
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
import xlrd
import xlwt
import os
import re
import NMRAMembersConfig
import DivisionMap
import pandas as pd

class NewsletterFile:
	#
	# Constructor
	#
	def __init__(self, roster_name, instance, work_dir, region, config, nmra_map):
		self.roster_name = roster_name
		self.output_name = roster_name
		self.work_dir = work_dir
		self.this_region = region
		self.config = config
		self.nmra_map = nmra_map
		self.roster_id = roster_name	# this will get reassigned if the output file is changed
		self.roster_ws = {}
		self.roster_ws_name = {}
		self.roster_ws_row = {}
		self.roster_ws_outputs = {}
		self.roster_ws_output_includes = {}
		self.roster_rs = None
		self.roster_rs_name = {}
		self.roster_nrows = 0
		self.roster_ncols = 0
		self.roster_includes = {}
		self.roster_sorts = {}
		self.header_index = {}
		self.instance = instance
		
		self.roster_ifmt = config.get_input_format()
		self.roster_ofmt = config.get_output_format()
		self.roster_outputs = config.get_outputs('newsletter', roster_name, self.instance)
		self.recipient_list = config.get_recipients('newsletter', roster_name, self.instance)
		self.recipients = {}
		
#		print("Recipients are: %s" % (self.recipient_list))
		if (len(self.roster_outputs) != 0):
			for output_name in self.roster_outputs:
				self.output_name = output_name
				self.roster_id = self.output_name
				self.roster_ws_name.update({self.roster_id : roster_name})
			pass
		else:
			self.roster_ws_name[self.roster_id] = roster_name
		pass

		fields = config.get_report_parameter_value('newsletter', roster_name, self.instance, 'fields').strip('"')
		format = config.get_report_parameter_value('newsletter', roster_name, self.instance, 'format').strip('"')
		self.roster_fields = fields.split(',')
		self.roster_format = format.split(',')
		self.roster_recipients = config.get_recipients('newsletter', roster_name, self.instance)
		includes = config.get_includes('newsletter', roster_name, self.instance)
		self.roster_sorts.update({self.roster_id : config.get_sorts('newsletter', roster_name, self.instance)})
		self.roster_includes.update({self.roster_id : includes})
		self.roster_ws.update({self.roster_id : pyexcel.Sheet()})
		self.roster_ws[self.roster_id].extend_rows(self.roster_format)
		self.roster_ws_row.update({self.roster_id : 1})

#		print("roster_name: %s(%d): " % (self.roster_name, self.instance))
#		print("output_name: %s: " % (self.output_name))
#		print("roster_fields: %s" % (self.roster_fields))
#		print("roster_format: %s" % (self.roster_format))
#		print("roster_recipients: %s" % (self.roster_recipients))
#		print("roster_includes: %s" % (self.roster_includes))
#		print("Instance: %d, %s" % (self.instance, self.roster_outputs))
	pass
	#
	# Read the given roster file and save the information necessary to create a new workbook from it.
	# This is where we are able to read in the crusty old version of Excel files.
	#
	def read_file(self, filename):
		#
		# Get the sheet
		#
		self.roster_rs = pyexcel.get_sheet(file_name=filename)
		self.roster_rs_name.update({self.roster_id : self.roster_rs.name})
		self.roster_nrows = self.roster_rs.number_of_rows()
		self.roster_ncols = self.roster_rs.number_of_columns()
		#
		# Figure out where all the data is in the header of the roster file
		#
		for field in self.roster_fields:
			col = 0
			found = False
#			print("%s:" % (field), end="")
			while (col < self.roster_ncols and not found):
				if (self.roster_rs.cell_value(0, col) == field):
					self.header_index.update({field : col})
					found = True
#					print("%s:" % (col), end="")
				else:
					col = col + 1
				pass
			pass
			if (not found):
				print("Error: Input field '%s' not found in %s!" % (field, filename))	
			pass		
		pass
		#
		# Find the column that has the requested include field in it
		# and store that include column number as the key to the list of
		# items given
		#
#		print("Looking for includes in %s..." % (self.roster_id))
		if (self.roster_id in self.roster_includes):
			value = self.roster_includes[self.output_name]
			if (len(value) > 0):
#				print("Include value = %s" % (value))
				vals = value.split('=')
				value_list = vals[1].split(',')
				values=[]
				for val in value_list:
					values.append(val)
				pass
#				print("Include values = %s" % (values))
				include_header = vals[0]
#				print("%s %s %s %s" % (key, value, values, include_header))
				hcol = self.header_index[include_header]
#				print("include_header = %s, hcol = %d" % (include_header, hcol))
				self.roster_ws_output_includes[self.roster_id] = {hcol : values}
			pass
		pass
	pass
	#
	#
	#
	def CustomLabelNameMapping(self, fname, mname, lname, org):
		r_fname = ""
		r_mname = ""
		r_lname = ""
		if (len(org) == 0):
			r_fname = fname
			r_mname = mname
			r_lname = lname
		else:
			orgs = re.split(',| ', org)
#			print("Converting: %s(%d)" % (orgs, len(orgs)))
			if (len(orgs) > 1):
#				print("'%s' '%s' '%s'" % (orgs[0:2], orgs[2:4], orgs[4:]))
				r_fname = ' '.join(orgs[0:2])
				if (len(orgs) > 3):
					r_mname = ' '.join(orgs[2:4])
				else:
					r_lname = orgs[2]
				pass
				if (len(orgs) > 4):
					r_lname = ' '.join(orgs[4:])
				pass
			else:
				r_fname = orgs[0]
				pass
#			print ("Returning '%s' '%s' '%s'" % (r_fname, r_mname, r_lname))
		pass
		return([r_fname, r_mname, r_lname])
	pass
			
	#
	# Copy member info from current worksheet to given output worksheet
	#
	# This code currently makes some assumptions about the mapping between the incoming roster
	# column order and the output order (self.roster_fields => self.roster_format). I don't have
	# time to code up the mapping in the configuration file for that now but that will have to 
	# be added at some point.
	#
	def write_row(self, row):
		skip = False
		if (self.roster_id in self.roster_ws_output_includes):
			for hcol in self.roster_ws_output_includes[self.roster_id]:
				values = self.roster_ws_output_includes[self.roster_id][hcol]
#				print("hcol = %s, values = %s" % (hcol, values))
				m_include = self.roster_rs.cell_value(row, hcol)
				if (not m_include in values):
					skip = True
#					print("Skipping member from %s" % (m_include))
				pass
			pass
		pass
		if (not skip):
			col = 0
			cols = len(self.roster_format)
			cells = []
			while (col <= cols-1):
				fmt = self.roster_format[col]
#				print("%d:%s, " % (col, fmt), end="")
				if (fmt == 'fname'):
					idx1 = self.header_index['fname']
					idx2 = self.header_index['mname']
					idx3 = self.header_index['lname']
					idx4 = self.header_index['Organization']
					val1 = self.roster_rs.cell_value(row, idx1)
					val2 = self.roster_rs.cell_value(row, idx2)
					val3 = self.roster_rs.cell_value(row, idx3)
					val4 = self.roster_rs.cell_value(row, idx4)
					val=[]
					val = self.CustomLabelNameMapping(val1, val2, val3, val4)
					cells.append(val[0])
					cells.append(val[1])
					cells.append(val[2])
					col = col + 3
				elif ('-' in fmt):
					zips = re.split(r'\-', fmt)
#					print("zips is %s => %s" % (fmt, zips))
					idx1 = self.header_index[zips[0]]
					idx2 = self.header_index[zips[1]]
					val1 = self.roster_rs.cell_value(row, idx1)
					val2 = self.roster_rs.cell_value(row, idx2)
					val = "%s-%s" % (val1, val2)
					cells.append(val)
					col = col + 1
				else:
					idx = self.header_index[fmt]
					val = self.roster_rs.cell_value(row, idx)
					cells.append(val)
					col = col + 1
				pass
			pass
#			print("")
#			print("Row: %s, Data: " % (self.roster_ws_row[self.roster_id]), end="")
#			print("%s" % (cells))
			self.roster_ws[self.roster_id].extend_rows(cells)
			self.roster_ws_row[self.roster_id] = self.roster_ws_row[self.roster_id] + 1
		pass
	pass
	#
	# Processes the current workbook in memory
	#
	def process(self, parent_dir):
		#
		# Write out the spreadsheets to all of the output files
		#
		for recipient in self.recipient_list:
			print("\t\tRecipient: %s" % (recipient))
			reg_fid = self.nmra_map.get_file_id(self.this_region, 0)
			reg_name = self.nmra_map.get_region_id(reg_fid)
			if (recipient == 'printer'):
				reg_dir = "%s/%s_Region_Printer" % (self.work_dir, reg_name)
			elif (recipient == 'editor'):
				reg_dir = "%s/%s_Region_Editor" % (self.work_dir, reg_name)
			else:
				reg_dir = "%s/%s_Region" % (self.work_dir, reg_name)
				
			os.makedirs("%s" % reg_dir, exist_ok=True)
			
			output_filename = "%s/%s.%s" % (reg_dir, self.output_name, self.roster_ofmt)
	
			
			for row in range(1, self.roster_nrows):
				self.write_row(row)
			pass
			if (len(self.roster_sorts[self.roster_id].keys()) > 0):
	#			print("We need to sort this one...")
				#
				# get the field from which we will sort the data
				#
				sort_list = self.config.get_sorts('newsletter', self.roster_name, self.instance)
				sort_field = sort_list.get('sort')
				#
				# write out the data to an Excel temp file so we can use Pandas on it
				#			
				temp_output_filename = "%s/%s.%s" % (reg_dir, self.output_name, 'xlsx')
				self.roster_ws[self.roster_id].name = self.roster_rs_name[self.roster_id]
				self.roster_ws[self.roster_id].save_as(temp_output_filename)
				#
				# Read it back into a panda dataframe
				#	
				excel_file = pd.ExcelFile(temp_output_filename)
				data_frame = excel_file.parse(self.roster_rs_name[self.roster_id])
				#
				# sort it using the sort field we found in the XML
				#
				data_frame.sort_values(by=sort_field, inplace=True, ignore_index=True)
				#
				# Now we have to drop the header row and the index column that sort added
				#
	#			print("%s" % (sort_frame.columns))
	#			sort_frame.at[0,''] = 'key'
	#			sort_frame.drop(sort_frame.columns[1], axis=1, inplace=True)	# column 0
	#			sort_frame.drop([0], inplace=True)			# row 0
				#
				# finally save it as a CSV file
				#
				print("\t\tOutput (sorted by '%s': %s" % (sort_field, output_filename))
				data_frame.to_csv(output_filename, header=None, index=False)
				if (recipient in self.recipient_list):
					if (not recipient in self.recipients.keys()):
						self.recipients.update({recipient: [output_filename]})
					else:
						self.recipients[recipient].append(output_filename)
					pass
				pass
				#
				# Remove the temporary Excel file
				#
				os.remove(temp_output_filename)
			else:
				#
				# at the end of the input sheet, write out all of the output sheets we made
				#
				print("\t\tOutput: %s" % (output_filename))
				self.roster_ws[self.roster_id].save_as(output_filename)
				if (recipient in self.recipient_list):
					if (not recipient in self.recipients.keys()):
						self.recipients.update({recipient: [output_filename]})
					else:
						self.recipients[recipient].append(output_filename)
					pass
				pass
			pass
		pass
		return(self.recipients)
	pass
pass

