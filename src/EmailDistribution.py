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
		self.required_distribution_columns = {'id' : -1, 'region' : -1, 'division' : -1, 'lname' : -1, 'fname' : -1, 'email_address' : -1, 'zip_file' : -1}
		self.settings_worksheet = "Email Settings"
		self.required_settings_columns = {'email_sender' : -1, 'email_bcc' : -1, 'email_smtp_address' : -1}
		
		self.sender_email = ""
		self.sender_bcc = ""
		self.smtp_address = ""
	pass
	#
	# Add a member to the distribution list
	#
	def add_recipient(self, nmra_id, region, division, lname, fname, email, zipfile):
		if nmra_id and (region != 0) and (division != 0) and lname and fname and email and zipfile:
			self.distribution_list.append({'nmra_id' : nmra_id, 'region' : region, 'division' : division, 'lname' : lname, 'fname' : fname, 'email' : email, 'zipfile' : zipfile, 'valid' : False})
		else:
			print("Warning: Recipient entry is incomplete, recipient not added.")
		pass
	pass
	#
	# Add a member to the distribution list
	#
	def validate_recipient(self, nmra_id, region, division, lname, fname, email):
		for recipient in self.distribution_list:
			if (recipient['nmra_id'] == nmra_id) and (recipient['region'] == region) and (recipient['division'] == division) and (recipient['lname'] == lname) and (recipient['fname'] == fname):
				if (recipient['email'] == email):
					recipient['valid'] = True
					print("\tEmail recipient ID %s (%s, %s) %02d%02d, email (%s) is valid" % (nmra_id, lname, fname, region, division, email))
				else:
					print("Warning: Email recipient ID %s (%s, %s) %02d%02d, their email (%s) doesn't match what the NMRA has (%s)" % (nmra_id, lname, fname, region, division, email, recipient['email']))
					recipient['valid'] = True
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
				r_zipfile  = "%s" % row[self.required_distribution_columns['zip_file']]
				if (r_id.startswith('#')):
					print("Skipping NMRA Member %s, (%s, %s) %02d%02d, email %s, zip file: %s" % (r_id, r_lname, r_fname, r_region, r_division, r_email, r_zipfile))
				else:
					print("NMRA Member %s, (%s, %s) %02d%02d email to %s with zip file: %s" % (r_id, r_lname, r_fname, r_region, r_division, r_email, r_zipfile))
					self.add_recipient(r_id, r_region, r_division, r_lname, r_fname, r_email, r_zipfile)
				pass
			pass
			row_num = row_num + 1	
		pass
	pass
	#
	# Send an email to the given recipient
	#
	def send_email(self, recipient, name, filepath, zip_file):
		#-------------------------------------------------------
		#
		# Pull out the required information
		#
		#-------------------------------------------------------
		if (recipient['zipfile'] == "NMRA"):
			filename	  = zip_file
			sendfile	  = zip_file
		else:
			filename	  = recipient['zipfile']
			sendfile	  = "%s/%s" % (filepath, filename)
		pass
		receiver_fname	  = recipient['fname']
		receiver_lname	  = recipient['lname']
		receiver_region   = recipient['region']
		receiver_division = recipient['division']
		sender_email	  = self.sender_email
		receiver_email	  = recipient['email']
		bcc_email		  = self.sender_bcc
		smtp_address	  = self.smtp_address
		print("\tSending NMRA %s %s to %s %s %02d%02d at %s..." % (name, filename, receiver_fname, receiver_lname, receiver_region, receiver_division, receiver_email))
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
	def send_emails(self, name, filepath, zip_file):
		for recipient in self.distribution_list:
			if (self.is_recipient_valid(recipient)):
				self.send_email(recipient, name, filepath, zip_file)
			else:
				print("Recipient NMRA ID %s (%s, %s) %02d%02d is not a valid NMRA Member! No email sent." % (recipient['nmra_id'], recipient['lname'], recipient['fname'], recipient['region'], recipient['division']))
			pass
		pass
	pass
pass

