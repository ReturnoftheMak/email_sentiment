# -*- coding: utf-8 -*-
"""
Created on Thu May 30 16:18:16 2019

@author: Makhan.Gill
"""

import csv
import sys
import json
csv_input_file = r'\\tcspmprf01\userdata$\Matthew.Rout\UserRedirect\Documents\MR_Emails2.CSV'

# Array of which types of emails to include
recipient_scope = ['to']

# Settings read into variables
settings_file = open(r'\\tcspmprf01\userdata$\Matthew.Rout\UserRedirect\Documents\email-extract\email-extract\settings.json', 'r')
settings = settings_file.read()
settings = json.loads(settings)

csv_keys = settings['csv_keys']
no_special_characters = settings['no_special_characters']
output_delimiter = settings['output_delimiter']
name_delimiter = settings['name_delimiter']
address_delimiter = settings['address_delimiter']
split_address_by_text = settings['split_address_by_text']
group_by_address_domain = settings['group_by_address_domain']
default_group_name = settings['default_group_name']
#address_string_to_ignore = settings['address_string_to_ignore']
address_string_to_ignore = '/O='
include_email_in_attributes = settings['include_email_in_attributes']

# Helper functions
def strip_non_ascii(string):
  # Returns the string without non ASCII characters
  stripped = (c for c in string if 0 < ord(c) < 127)
  return ''.join(stripped)

def process_name(string):
  string = string.strip().replace("'","")
  if string is None or string == "":
    string = "BLANK"
  if no_special_characters:
    string = strip_non_ascii(string)
  return string

# Process update
print("")
print("Analyzing " + "'" + csv_input_file + "'...")
print( "")

# Prompts for what to include
#include_cc = raw_input("Include CC recipients (Y/N): ").upper()
include_cc = 'N'
print("")

if include_cc == 'Y':
  recipient_scope.append('cc')
  include_bcc = raw_input("Include BCC recipients (Y/N): ").upper()
  if include_bcc == 'Y':
    recipient_scope.append('bcc')
  print("")

max_recipients = 0
#trigger_max_recipients = raw_input("Only consider emails sent out to a limited number of people (Y/N): ").upper()
trigger_max_recipients = 'Y'
print ("")
'''
if trigger_max_recipients == 'Y':
  max_recipients = raw_input("Enter max. number of recipients: ")
  print ""
else:
  trigger_max_recipients = 'N'
'''
max_recipients = 20

# Open input file
input_file = csv.DictReader(open(csv_input_file))

# Variables to store node and tie details
nodes = {}
nodes_with_ids = {}
ties = {}

# Preparing text strings
group_text = "\"sep=" + output_delimiter + "\"" + "\n"
group_text += "group_id" + output_delimiter + "group_name\n"
ties_text = "\"sep=" + output_delimiter + "\"" + "\n"
ties_text += "sender" + output_delimiter + "recipient" + output_delimiter + "count\n"
attribute_text = "\"sep=" + output_delimiter + "\"" + "\n"
attribute_text += "\"id\"" + output_delimiter + "\"name\"" + output_delimiter + "\"group_id\""
if include_email_in_attributes:
  attribute_text += output_delimiter + "\"email\""
attribute_text += "\n"


# Read each line in the CSV file
for row in input_file:
        
  # Iterate through each scope that has been included (to, cc, bcc)  
  for scope in recipient_scope:
      
    # Populating nodes with recipients
    recipient_addresses = []
    sender_addresses = []

    if row[csv_keys[scope + '_email']] is not None and row[csv_keys[scope + '_name']] is not None:
      temp_recipient_addresses = row[csv_keys[scope + '_email']].upper().replace(address_delimiter, split_address_by_text).split(split_address_by_text)
      recipient_names = row[csv_keys[scope + '_name']].upper().split(name_delimiter)

      for temp_recipient_address in temp_recipient_addresses:
        if address_string_to_ignore not in temp_recipient_address:
          recipient_addresses.append(temp_recipient_address)

      if len(recipient_addresses) > len(recipient_names):
        recipient_addresses.pop(0)
      elif len(recipient_addresses) != len(recipient_names):
        print('PROBLEM: ' + str(len(recipient_addresses)) + " ADDRESSES VS " + str(len(recipient_names)) + " NAMES")

    # Check against max count and skip if more included
    if trigger_max_recipients == 'N' or (trigger_max_recipients == 'Y' and len(recipient_names) <= int(max_recipients)):
      for idx, recipient_address in enumerate(recipient_addresses):
        nodes[recipient_address.split(address_delimiter)[0]] = recipient_names[idx]

      # Populating nodes with senders
      if row[csv_keys['from_email']] is not None and row[csv_keys['from_name']] is not None:
        temp_sender_addresses = row[csv_keys['from_email']].upper().replace(address_delimiter, split_address_by_text).split(split_address_by_text)
        sender_names = row[csv_keys['from_name']].upper().split(";")

        for temp_sender_address in temp_sender_addresses:
          if address_string_to_ignore not in temp_sender_address:
            sender_addresses.append(temp_sender_address)

      # Populating ties
      for idx, sender_address in enumerate(sender_addresses):
     
        nodes[sender_address.upper().split(address_delimiter)[0]] = sender_names[idx]
  
        if sender_address not in ties:
          ties[sender_address] = {}
        for recipient_address in recipient_addresses:
          recipient_address = recipient_address.upper().split(";")[0]
          if recipient_address not in ties[sender_address]:
            ties[sender_address][recipient_address] = 0
          ties[sender_address][recipient_address] += 1
    
# Variables for setting ids and capture group details
node_id = 1
groups = {}
group_id = 1
nodes_listed_by_id = {}

# Structure node information and write group text
for node_email, node_name in nodes.items():
  nodes_with_ids[node_email] = {}
  nodes_with_ids[node_email]['id'] = node_id
  nodes_with_ids[node_email]['name'] = node_name
  nodes_with_ids[node_email]['email'] = node_email
  nodes_listed_by_id[node_id] = node_email

  # Capturing group name if relevant - TODO: also if not set
  if group_by_address_domain:
    split_email = node_email.split("@")
    if len(split_email) > 1:
      group_name = split_email[1].strip()
    else:
      group_name = default_group_name
    if group_name not in groups.keys():
      groups[group_name] = group_id
      group_id += 1
  else:
    # Note: will be set to the same over an over again
    group_name = default_group_name
    groups[group_name] = group_id 

  nodes_with_ids[node_email]['group_name'] = group_name
  nodes_with_ids[node_email]['group_id'] = groups[group_name]
  node_id += 1

# Write group text
for group_name, group_id in groups.items():
  group_text += str(group_id) + output_delimiter + "\"" + group_name + "\"\n"

# Write tie text
for sender, recipients in ties.items():
  for recipient, count in recipients.items():
    ties_text += str(nodes_with_ids[sender]['id']) + output_delimiter + str(nodes_with_ids[recipient]['id']) + output_delimiter + str(count) + "\n"
    
# Write attribute text
for node_id, node_email in nodes_listed_by_id.items():
  node_details = nodes_with_ids[node_email]

  attribute_text += str(node_details['id']) + output_delimiter + "\"" + process_name(node_details['name']) + "\"" + output_delimiter
  if group_by_address_domain:
    attribute_text += str(node_details['group_id'])
  else:
    attribute_text += str(1)
  if include_email_in_attributes:
    attribute_text += output_delimiter + "\""  + process_name(node_details['email'])  + "\""
  attribute_text += "\n"

# Progress update
print ("Analysis completed - see the generated CSV files.")
print ("")

# Write output files
#attribute_file = open('attributes.csv', 'w')
#attribute_file.write(attribute_text)
#attribute_file.close()

#groups_file = open('groups.csv', 'w')
#groups_file.write(group_text)
#groups_file.close()

#email_counts_file = open('email_counts.csv', 'w')
#email_counts_file.write(ties_text)
#email_counts_file.close()