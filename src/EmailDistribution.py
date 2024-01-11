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
		self.required_distribution_columns = {'id' : -1, 'region' : -1, 'division' : -1, 'lname' : -1, 'fname' : -1, 'email_address' : -1, 'category' : -1, 'file' : -1, 'notes' : -1}
		self.settings_worksheet = "Email Settings"
		self.required_settings_columns = {'email_sender' : -1, 'email_bcc' : -1, 'email_smtp_address' : -1}
		
		self.sender_email = ""
		self.sender_bcc = ""
		self.smtp_address = ""
	pass
	#
	# Add a member to the distribution list
	#
	def add_recipient(self, nmra_id, region, division, lname, category, fname, email, file, notes):
		if nmra_id and (region != 0) and (division != 0) and lname and fname and email and category:
			self.distribution_list.append({'nmra_id' : nmra_id, 'region' : region, 'division' : division, 'lname' : lname, 'fname' : fname, 'email' : email, 'category' : category, 'file' : file, 'notes' : notes, 'zip_file' : "n/a", 'location' : "Unknown Division", 'valid_member' : False, 'valid_email' : False})
		else:
			print("Warning: Recipient entry is incomplete, recipient not added.")
		pass
	pass
	#
	# Add a member to the distribution list
	#
	def validate_recipient(self, nmra_map, parent_dir, dist_dir, zip_filename, nmra_id, region, division, lname, fname, email, force_override):
		for x in range(0, len(self.distribution_list)):
			try:
				reg_id = self.distribution_list[x].get('region')
			except ValueError:
				print("Unknown region error in the email distribution list")
			pass
			try:
				div_id = self.distribution_list[x].get('division')
			except ValueError:
				print ("Unknown division error in the email distribution list")
			pass

			reg_fid = nmra_map.get_file_id(reg_id, 0)
			div_fid = nmra_map.get_file_id(reg_id, div_id)

			reg_name = nmra_map.get_region(reg_fid)
			div_name = nmra_map.get_division(div_fid)
			
			list_div_name = nmra_map.get_division(nmra_map.get_file_id(reg_id, division))

			if (self.distribution_list[x].get('nmra_id') == nmra_id) and (self.distribution_list[x].get('region') == region) and (self.distribution_list[x].get('lname') == lname) and (self.distribution_list[x].get('fname') == fname):
				location = "%s %s Division" % (reg_name, div_name)
				self.distribution_list[x].update({'location' : location})
				self.distribution_list[x].update({'valid_member' : True})
				if (((self.distribution_list[x].get('email') == email) and (self.distribution_list[x].get('division') == division)) or (force_override and ((self.distribution_list[x].get('email') != email) or (self.distribution_list[x].get('division') != division)))):
					self.distribution_list[x].update({'valid_email' : True})
					#
					# NMRA = zip file of the entire directory of the processed results
					# REGION = just the zip file of the region informaiton
					# DIVISION = just the zip file of the divsion information
					# filename.zip = explicitly send the zip file with the name filename.zip in the release directory
					#
					if (self.distribution_list[x].get('category') == "NMRA"):
						zip_file = "%s/%s/../%s_processed.zip" % (parent_dir, dist_dir, zip_filename)
					elif (self.distribution_list[x].get('category') == "REGION"):
						zip_file = "%s/%s/%s_Region.zip" % (parent_dir, dist_dir, reg_name)
					elif (self.distribution_list[x].get('category') == "DIVISION"):
						zip_file = "%s/%s/%s_Region-%s_Division.zip" % (parent_dir, dist_dir, reg_name, div_name)
					elif (self.distribution_list[x].get('category') == "PRINTER"):
						zip_file = "%s/%s/%s_Region_Printer.zip" % (parent_dir, dist_dir, reg_name)
					elif (self.distribution_list[x].get('category') == "EDITOR"):
						zip_file = "%s/%s/%s_Region_Editor.zip" % (parent_dir, dist_dir, reg_name)
					elif (self.distribution_list[x].get('category') == "FILE"):
						zip_file = "%s/%s/%s" % (parent_dir, dist_dir, self.distribution_list[x].get('file'))
					pass
					self.distribution_list[x].update({'zip_file' : zip_file})
					if (force_override and ((self.distribution_list[x].get('email') != email) or (self.distribution_list[x].get('division') != division))):
						print("\tWarning: Email recipient %s %s, NMRA ID: %s, from %s %s Division, their email (%s) doesn't match what the NMRA has (%s) for %s %s Division" % (fname, lname, nmra_id, reg_name, list_div_name, self.distribution_list[x].get('email'), email, reg_name, div_name))
					else:
						short_zip_file = os.path.basename(zip_file)
						print("\tValidated Email recipient %s %s, NMRA ID: %s, from %s %s Division, email (%s) will receive %s" % (fname, lname, nmra_id, reg_name, div_name, email, short_zip_file))
					pass
				else:
					print("\tWarning: Email recipient %s %s, NMRA ID: %s, from %s %s Division, their email (%s) doesn't match what the NMRA has (%s) for %s %s Division" % (fname, lname, nmra_id, reg_name, div_name, self.distribution_list[x].get('email'), email, reg_name, div_name))
					print("\tWarning: NOT sending email to this recipient, use the -f option to override unrecognized email addresses!")
					self.distribution_list[x].update({'valid_email' : False})
				pass
			pass
		pass
	pass
	#
	# is a recipient a valid NMRA member
	#
	def is_recipient_valid_member(self, recipient):
		if (recipient.get('valid_member')):
			return True
		else:
			return False
		pass
	pass
	#
	# does recipient have a valid NMRA email
	#
	def is_recipient_valid_email(self, recipient):
		if (recipient.get('valid_email')):
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
			settings_ws = distribution_wb[	self.settings_worksheet]
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
				if (r_id.startswith('#')):
					print("Skipping comment: %s" % (row))
				else:
					r_region   =    int(row[self.required_distribution_columns['region']])
					r_division =    int(row[self.required_distribution_columns['division']])
					r_lname    = "%s" % row[self.required_distribution_columns['lname']]
					r_fname    = "%s" % row[self.required_distribution_columns['fname']]
					r_email    = "%s" % row[self.required_distribution_columns['email_address']]
					r_category = "%s" % row[self.required_distribution_columns['category']]
					r_file 	   = "%s" % row[self.required_distribution_columns['file']]
					r_notes	   = "%s" % row[self.required_distribution_columns['notes']]
					print("NMRA Member %s %s (ID %s) email to %s category %s with file: %s, notes: %s" % (r_fname, r_lname, r_id, r_email, r_category, r_file, r_notes))
					self.add_recipient(r_id, r_region, r_division, r_lname, r_category, r_fname, r_email, r_file, r_notes)
				pass
			pass
			row_num = row_num + 1	
		pass
	pass
	#
	# print email list
	#
	def print_email_list(self):
		print("Send the following emails:")
		print("")
		for recipient in self.distribution_list:
			sendfile	      = recipient.get('zip_file')
			filename	  	  = os.path.basename(sendfile)
			receiver_fname	  = recipient.get('fname')
			receiver_lname	  = recipient.get('lname')
			receiver_email	  = recipient.get('email')
			valid_email		  = recipient.get('valid_email')
			bcc_email		  = self.sender_bcc
			if (valid_email):
				valid_string = 'Y'
			else:
				valid_string = 'N'
			pass
			print("%-30s (%s) %-30s %s" % ("%s %s" % (receiver_fname, receiver_lname), valid_string, receiver_email, filename))
#			print("%s" % recipient)	
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
		sendfile	      = recipient.get('zip_file')
		filename	  	  = os.path.basename(sendfile)
		receiver_fname	  = recipient.get('fname')
		receiver_lname	  = recipient.get('lname')
		receiver_location = recipient.get('location')
		receiver_notes	  = recipient.get('notes')
		sender_email	  = self.sender_email
		receiver_email	  = recipient.get('email')
		bcc_email		  = self.sender_bcc
		smtp_address	  = self.smtp_address
		print("\tSending NMRA Roster from %s %s to %s %s from %s, to email %s..." % (name, filename, receiver_fname, receiver_lname, receiver_location, receiver_email))
		#-------------------------------------------------------
		#
		# Create the message
		#
		#-------------------------------------------------------
		subject = "NMRA Roster File (%s): %s" % (name, filename)
		body = 		  "Hello %s,\n\nPlease find the attached NMRA roster file %s from the NMRA distribution for %s\n\n" % (receiver_fname, filename, name)
		body = body + "You are receiving this NMRA roster for: %s\n\n" % (receiver_notes)
		body = body + "Please reply to this message if you are no longer the proper recipient of this information.\n\n"
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
			if (self.is_recipient_valid_member(recipient)):
				if (self.is_recipient_valid_email(recipient)):
					self.send_email(recipient, name)
				else:
					print("\tWarning: Recipient NMRA %s %s (ID %s) %s does not have a valid NMRA EMail address! Their email address does not match what is on file!" % (recipient.get('fname'), recipient.get('lname'), recipient.get('nmra_id'), recipient.get('location')))
				pass
			else:
				print("\tWarning: Recipient NMRA %s %s (ID %s) %s is not a valid NMRA Member! They were not found in %s.zip. No email sent." % (recipient.get('fname'), recipient.get('lname'), recipient.get('nmra_id'), recipient.get('location'), name))
			pass
		pass
	pass
pass

