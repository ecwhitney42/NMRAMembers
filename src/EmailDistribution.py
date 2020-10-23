#------------------------------------------------------------------------------
#
# Class EmailDistribution
#
# The file format for the email distribution file is a .xlsx file with two worksheets:
#
# Email Settings Worksheet
# email_sender, email_bcc, email_smtp_address

# Email Distribution Worksheet
# id, region, division, lname, fname, email_address, zip_file
#
# id:				the NMRA number of the recipient
# region:			the two-digit NMRA region number of the recipient
# division:			the two-digit NMRA division number of the recipient
# lname:			the last name of the recipient
# fname:			the first name of the recipient
# email_address:	the email address recipient
# zip_file:			the name of the zip file to send
#------------------------------------------------------------------------------
#
# This class manages the email distribution list
#
#------------------------------------------------------------------------------
import sys
import os
import pyexcel
import pyexcel_xlsx
import io
import email, smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class EmailDistribution:
	#
	# Constructor
	#
	def __init__(self):
		self.distribution_list = []
		self.distribution_worksheet = "Email Distribution"
		self.required_distribution_columns = {'id' : -1, 'region' : -1, 'division' : -1, 'lname' : -1, 'fname' : -1, 'email_address' : -1, 'file' : -1}
		self.settings_worksheet = "Email Settings"
		self.required_settings_columns = {'email_sender' : -1, 'email_bcc' : -1, 'email_smtp_address' : -1}
		
		self.sender_email = ""
		self.sender_bcc = ""
		self.smtp_address = ""
	pass
	#
	# Add a member to the distribution list
	#
	def add_recipient(self, nmra_id, region, division, lname, fname, email, file):
		if nmra_id and (region != 0) and (division != 0) and lname and fname and email and file:
			self.distribution_list.append({'nmra_id' : nmra_id, 'region' : region, 'division' : division, 'lname' : lname, 'fname' : fname, 'email' : email, 'file' : file, 'zip_file' : "n/a", 'location' : "Unknown Division", 'valid' : False})
		else:
			print("Warning: Recipient entry is incomplete, recipient not added.")
		pass
	pass
	#
	# Add a member to the distribution list
	#
	def validate_recipient(self, use_long, nmra_map, parent_dir, dist_dir, zip_filename, nmra_id, region, division, lname, fname, email, force_override):
		for recipient in self.distribution_list:
			if (recipient['nmra_id'] == nmra_id) and (recipient['region'] == region) and (recipient['division'] == division) and (recipient['lname'] == lname) and (recipient['fname'] == fname):
				if (recipient['email'] == email):
					recipient['valid'] = True
					reg_id = recipient['region']
					div_id = recipient['division']

					reg_fid = nmra_map.get_file_id(reg_id, 0)
					div_fid = nmra_map.get_file_id(reg_id, div_id)

					if use_long:
						reg_name = nmra_map.get_region(reg_fid)
					else:
						reg_name = nmra_map.get_region_id(reg_fid)
					pass
					div_name = nmra_map.get_division(div_fid)
					#
					# NMRA = zip file of the entire directory of the processed results
					# REGION = just the zip file of the region informaiton
					# DIVISION = just the zip file of the divsion information
					# filename.zip = explicitly send the zip file with the name filename.zip in the release directory
					#
					if (recipient['file'] == "NMRA"):
						zip_file = "%s/%s/../%s_processed.zip" % (parent_dir, dist_dir, zip_filename)
					elif (recipient['file'] == "REGION"):
						zip_file = "%s/%s/%s_Region.zip" % (parent_dir, dist_dir, reg_name)
					elif (recipient['file'] == "DIVISION"):
						zip_file = "%s/%s/%s_Region-%s_Division.zip" % (parent_dir, dist_dir, reg_name, div_name)
					else:
						zip_file = "%s/%s/%s" % (parent_dir, dist_dir, recipient['file'])
					pass
					recipient['zip_file'] = zip_file
					if (not reg_name or not div_name):
						location = "Unknown Division"
					else:
						location = "%s %s Division" % (reg_name, div_name)
					pass
					recipient['location'] = location
					print("\tEmail recipient %s %s, NMRA ID: %s, from %s %s Division, email (%s) will receive %s" % (fname, lname, nmra_id, reg_name, div_name, email, zip_file))
				else:
					print("Warning: Email recipient %s %s, NMRA ID: %s, from %s %s Division, their email (%s) doesn't match what the NMRA has (%s)" % (fname, lname, nmra_id, reg_name, div_name, email, recipient['email']))
					if (force_override):
						print("Warning: Forcing override of email address given in the config file: %s" % (email))
						recipient['valid'] = True
					else:
						print("Warning: NOT sending email to this recipient!")
						recipient['valid'] = False
					pass
				pass
			pass
		pass
	pass
	#
	# is a recipient a valid NMRA member
	#
	def is_recipient_valid(self, recipient):
		if (recipient['valid']):
			return True
		else:
			return False
		pass
	pass
	#
	# Reads the reassignment file and populates its data structure in memory
	#
	def read_file(self, filename):
		print("Reading the NMRA Email Distribution File: %s" % filename)
		try:
			distribution_wb = pyexcel.get_book(file_name=filename)
			distribution_ws = distribution_wb[self.distribution_worksheet]
			settings_ws = distribution_wb[self.settings_worksheet]
		except:
			print("Email Distribution Spreadsheet Error: ", sys.exc_info()[0])
			raise
		pass
		#
		# Process the email settings worksheet
		#
		all_good = True
		row_num = 0
		for row in settings_ws:
			#
			# The first row contains the column headings we need to find the offsets
			#
			if (row_num == 0):
				col_num = 0
				for cell in row:
					for key in self.required_settings_columns.keys():
						if (cell == key):
							self.required_settings_columns[key] = col_num
						pass
					col_num = col_num + 1
				pass
				for key in self.required_settings_columns.keys():
					if (self.required_settings_columns[key] == -1):
						all_good = False
					pass
				pass
				if (all_good == False):
					raise ValueError('All required columns MUST be included in the Email Settings Worksheet!')
				pass
			#
			# The only settings that matter are the ones in the row right below the headers
			#
			elif (row_num == 1):
				self.sender_email  = "%s" % row[self.required_settings_columns['email_sender']]
				self.sender_bcc    = "%s" % row[self.required_settings_columns['email_bcc']]
				self.smtp_address  = "%s" % row[self.required_settings_columns['email_smtp_address']]
				#
				# Report the settings found
				#
				print("The Emails will be sent by %s, BCC'd to %s, SMTP = %s" % (self.sender_email, self.sender_bcc, self.smtp_address))
			pass
			row_num = row_num + 1	
		pass
		#
		# Process the email distribution list worksheet
		#
		all_good = True
		row_num = 0
		for row in distribution_ws:
			#
			# The first row contains the column headings we need to find the offsets
			#
			if (row_num == 0):
				col_num = 0
				for cell in row:
					for key in self.required_distribution_columns.keys():
						if (cell == key):
							self.required_distribution_columns[key] = col_num
						pass
					col_num = col_num + 1
				pass
				for key in self.required_distribution_columns.keys():
					if (self.required_distribution_columns[key] == -1):
						all_good = False
					pass
				pass
				if (all_good == False):
					raise ValueError('All required columns MUST be included in the Email Distribution Worksheet!')
				pass
			#
			# All subsequent rows contain the recipient data
			#
			else:
				r_id       = "%s" % row[self.required_distribution_columns['id']]
				r_region   =    int(row[self.required_distribution_columns['region']])
				r_division =    int(row[self.required_distribution_columns['division']])
				r_lname    = "%s" % row[self.required_distribution_columns['lname']]
				r_fname    = "%s" % row[self.required_distribution_columns['fname']]
				r_email    = "%s" % row[self.required_distribution_columns['email_address']]
				r_file 	   = "%s" % row[self.required_distribution_columns['file']]
				if (r_id.startswith('#')):
					print("Skipping NMRA Member %s %s (ID %s) email %s, file: %s" % (r_fname, r_lname, r_id, r_email, r_file))
				else:
					print("NMRA Member %s %s (ID %s) email to %s with file: %s" % (r_fname, r_lname, r_id, r_email, r_file))
					self.add_recipient(r_id, r_region, r_division, r_lname, r_fname, r_email, r_file)
				pass
			pass
			row_num = row_num + 1	
		pass
	pass
	#
	# Send an email to the given recipient
	#
	def send_email(self, recipient, name):
		#-------------------------------------------------------
		#
		# Pull out the required information
		#
		sendfile	      = recipient['zip_file']
		filename	  	  = os.path.basename(sendfile)
		receiver_fname	  = recipient['fname']
		receiver_lname	  = recipient['lname']
		receiver_location = recipient['location']
		sender_email	  = self.sender_email
		receiver_email	  = recipient['email']
		bcc_email		  = self.sender_bcc
		smtp_address	  = self.smtp_address
		print("\tSending NMRA Roster from %s %s to %s %s from %s, to email %s..." % (name, filename, receiver_fname, receiver_lname, receiver_location, receiver_email))
		#-------------------------------------------------------
		#
		# Create the message
		#
		#-------------------------------------------------------
		subject = "NMRA Roster File (%s): %s" % (name, filename)
		body = "Hello %s,\n\nPlease find the attached NMRA roster file %s from the NMRA distribution for %s\n\nPlease reply to this message if you are no longer the proper recipient of this information.\n\n" % (receiver_fname, filename, name)
		#-------------------------------------------------------
		#
		# Create a multipart message and set headers
		#
		#-------------------------------------------------------
		message = MIMEMultipart()
		message["From"]    = sender_email
		message["To"]      = receiver_email
		message["Subject"] = subject
		if not bcc_email.strip():
			message["Bcc"]     = bcc_email  # Recommended for mass emails
		pass
		#-------------------------------------------------------
		#
		# Add body to email and put it together
		#
		#-------------------------------------------------------
		message.attach(MIMEText(body, "plain"))

		# Open PDF file in binary mode
		with open(sendfile, "rb") as attachment:
			# Add file as application/octet-stream
			# Email client can usually download this automatically as attachment
			part = MIMEBase("application", "zip")
			part.set_payload(attachment.read())
		pass
		#
		# Encode file in ASCII characters to send by email    
		#
		encoders.encode_base64(part)
		#
		# Add header as key/value pair to attachment part
		#
		part.add_header("Content-Disposition", "attachment; filename= %s" % filename)
		#
		# Add attachment to message and convert message to string
		#
		message.attach(part)
		text = message.as_string()
		#-------------------------------------------------------
		#
		# Try to send the email
		#
		#-------------------------------------------------------
		try:	
			smtpObj = smtplib.SMTP(smtp_address)      
			smtpObj.sendmail(sender_email, receiver_email, text)        
		except smtplib.SMTPException:                          
			print("\tError: unable to send email")                
		pass
	pass
	#
	# Send an email to the given recipient
	#
	def send_emails(self, name):
		
		for recipient in self.distribution_list:
			if (self.is_recipient_valid(recipient)):
				self.send_email(recipient, name)
			else:
				print("\tWarning: Recipient NMRA %s %s (ID %s) %s is not a valid NMRA Member! They were not found in %s.zip. No email sent." % (recipient['fname'], recipient['lname'], recipient['nmra_id'], recipient['location'], name))
			pass
		pass
	pass
pass

