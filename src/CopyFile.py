#------------------------------------------------------------------------------
#
# Class CopyFile
#
# Copy a roster file
#
#------------------------------------------------------------------------------
import sys
import pandas as pd 
import os
import re
import NMRAMembersConfig
import DivisionMap

class CopyFile:
	#
	# Constructor
	#
	
	def __init__(self):
		return
	pass

	def process(self, roster_file, instance, roster_filename, work_dir, oformat, region, config, nmra_map):
		recipient_list = config.get_recipients('copy', roster_file, instance)
		recipients = {}
		for recipient in recipient_list:
			self.recipients.update({recipient : []})
		pass 
		
		report_file = os.path.splitext(os.path.basename(roster_file))
		report_name = report_file[0]
		re1 = re.compile(r'^(\d+)_(.*)')
		report_name = re1.sub(r'\2', report_name)
		re2 = re.compile(r'(.*)(\d+)$')
		report_name = re2.sub(r'\1', report_name)
			
		#
		# Get the workbook
		#
		roster_exf = pd.ExcelFile(roster_filename)
		roster_rdf = roster_exf.parse()
		#
		#
		# Make all of the column headings lower case
		#
		for col in range(0, len(roster_rdf.columns)):
			old = roster_rdf.columns[col]
			new = old.lower()
			if (new == 'memberregion'):
				region_header = new
			pass
			roster_rdf.rename(columns={old : new}, inplace=True)
		pass
		#
		# Fix the date field
		#
		for field in date_fields:
			roster_rdf[field]=pd.to_datetime(roster_rdf[field])
			roster_rdf[field]=roster_rdf[field].dt.strftime(date_format)
		pass
		#
		# create the output directory
		#
		reg_fid = nmra_map.get_file_id(region, 0)
		reg_name = nmra_map.get_region_id(reg_fid)
		#
		# generate the output filename
		#	
		for recipient in recipient_list:
			if (recipient == 'region'):
				reg_dir = "%s/%s_Region" % (work_dir, reg_name)
			elif (recipient == 'editor'):
				reg_dir = "%s/%s_Region_Editor" % (work_dir, reg_name)
			pass
			os.makedirs("%s" % reg_dir, exist_ok=True)
			outbasename = os.path.basename(roster_file)
			savepath = "%s/%s.%s" % (reg_dir, outbasename, oformat)
			#
			# save the file
			#
			roster_rdf.to_csv(savepath, index=False, date_format=date_format)
			if (recipient in recipient_list):
				filename = "%s/%s.%s" % (reg_dir, report_name, oformat)
				if (not recipient in recipients.keys()):
					recipients.update({recipient : [filename]})
				else:
					recipients[recipient].append(filename)
				pass
			pass
		return(recipients)
		pass
	pass
pass

