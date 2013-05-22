#!/usr/bin/python

import os, sys
import poplib
import email
import datetime

# This is that path that this script resides in.
# Don't edit this
script_dir = os.path.dirname( os.path.abspath(sys.argv[0]) )

# Settings file location
# This file should exist alongside this script
settings_file = os.path.join(script_dir, 'process_webfone_emails.conf')

# By default, we're assuming that the script is in the base directory
#
# You may specify another path in the settings file.  If you do, this
#	will be overwritten a few lines down after we open the settings file.
base_dir = os.path.join(script_dir, '[user]')

# Get user information from the settings file
user_list = {}

FILE = open(settings_file, 'r')
lines = FILE.readlines()
FILE.close()

for line in lines:
	if ( line[:9] == 'base_dir=' ):
		base_dir = line.rstrip().split('base_dir=', 1)[-1].replace('[script_dir]', script_dir)
	
	if ( line[:5] == 'user=' ):
		i = lines.index(line)
		u = line.rstrip().split('user=')[-1]
		a = lines[i+1].rstrip().split('=', 1)
		b = lines[i+2].rstrip().split('=', 1)
		user_list[u] = {
			a[0]: a[1],
			b[0]: b[1]
		}

########################################################################

# Connect to the POP3 server
def connect_email(user, passw, pop_url):
	print('Connecting to %s ...' % user)
	try:
		conn = poplib.POP3_SSL(pop_url)
		conn.user(user)
		conn.pass_(passw)
	except:
		return False, 'Could not connect to %s at %s' % (user, pop_url)
		
	return True, conn

# Return a list of messages
def get_messages(conn):
	print('Retrieving message list ...')
	msg_list = conn.list()
	return msg_list

# Retrieve a message
def retrieve_message(conn, msg):
	num, octet = msg.split(' ')
	print('Checking message %s' % str(num) )

	msg_contents = conn.retr(num)
	msg_details = email.message_from_string('\n'.join(msg_contents[1]))
	return msg_details

# Find the attachment
def get_attachment(msg_details):
	payload = msg_details.get_payload()
	if ( len(payload) < 2 ):
		print('No attachment found.')
		return False
	else:
		print('Found %s' % payload[1].get_filename())
		return payload[1]

# See if a folder exists
def check_dir(folder):
	if not os.path.exists(folder):
		try:
			print('Creating folder %s ...' % folder)
			os.mkdir(folder)
		except:
			return False, 'Could not make directory %s' % folder
	
	return True, ''

# Write the attachment to the desired folder
def write_attachment(user, attachment, msg_date):
	
	# Make a folder name based off the user
	folder = base_dir.replace('[user]', user)
	
	# Parse the year/month/day from the message so we can put it
	#	in the appropriate folder
	msg_timestamp = datetime.datetime.strptime(' '.join(msg_date.split(' ')[:-1]), '%a, %d %b %Y %H:%M:%S')
	year_folder = os.path.join( folder, str(msg_timestamp.year) )
	month_folder = os.path.join( year_folder, str(msg_timestamp.month) )
	day_folder = os.path.join( month_folder, str(msg_timestamp.day) )
	
	# Check to make sure we have the folders we need
	folder_list = [folder, year_folder, month_folder, day_folder]
	for f in folder_list:
		succ, mess = check_dir(f)
		if not succ: return False, mess
	
	# Get the attachment contents
	file_contents = attachment.get_payload(decode=True)
	
	# Build a full path for the file
	file_name = attachment.get_filename()
	file_path = os.path.join(day_folder, file_name)

	# Try to write the file.
	print('Saving %s ...' % file_path)
	try:
		with open(file_path, 'wb') as FILE:
			FILE.write(file_contents)
	except: return False, 'Could not write attachment'
	return True, 'Saved %s' % file_path

########################################################################

if __name__ == '__main__':
	
	# Do this for every user we know from the option above
	for user in user_list:
		
		# Print a newline to help readability
		print('')
		
		# Get the information we need
		pop_url = user_list[user]['pop_url']
		passw = user_list[user]['passw']
		
		# Connect and retrieve message list
		succ, conn = connect_email(user, passw, pop_url)
		if not succ:
			print(conn)
			continue
			
		msg_list = get_messages(conn)
		
		# If there are no new messages, carry on
		if ( len(msg_list[1]) == 0 ):
			print('No new messages found.')
			continue
		
		# Loop through for each message found
		for msg in msg_list[1]:

			# Print a newline to help readability
			print('')
		
			msg_details = retrieve_message(conn, msg)
			attachment = get_attachment(msg_details)
			
			if not attachment: continue
			else:
				succ, mess = write_attachment(user, attachment, msg_details['Date'])
				if not succ:
					print(mess)
					continue
		
		conn.quit()
	
	print('')
	sys.exit(0)
