# -*- coding: utf-8 -*-
"""
Created on Wed May 29 11:00:06 2019

@author: Makhan.Gill
"""

# Standard Imports

import win32com.client
import pandas as pd
import numpy as np
import datetime


#%% Select inbox folder


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

Inbox = outlook.GetDefaultFolder(6)

Sent = outlook.GetDefaultFolder(5)

# "6" refers to the index of a folder - in this case,
# the inbox. You can change that number to reference any other folder


#%% Test for content

messages = Inbox.Items
#message = messages.GetFirst()
#body_content = message.body
#print(body_content)


#%% This converts the outlook time to a pandas interpretable datetime

def convert_to_datetime(windows_time):
    """ Convert the pywintypes.datetime to a pandas datetime format.
    Note that instead of just stripping out the timezone it may need to be utilised if this is to become usable across timezones
    
    Args:
        arg1 windows_time (pywintypes.datetime): Timestamp from outlook item
        
    Returns:
        datetime
    """
    
    for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%d %H:%M'):
        try:
            datetime_py = datetime.datetime.strptime(str(windows_time).rstrip("+00:00"),fmt)
            return datetime_py
        except ValueError:
            pass


#%% Create Dictionary from message.Recipients

def create_recipients(message):
    """ Use to, cc and bcc to create recipients with error handling
    
    Args:
        arg1 message (win32com.client.CDispatch):
      
    Returns
        Dict
    """
    
    try:
        if ';' in message.to:
            a = message.to.split(";")
        else:
            a = [message.to]
    except:
        a = []
    
    try:
        if ';' in message.CC:
            b = message.CC.split(";")
        else:
            b = [message.CC]
    except:
        b = []
        
    try:
        if ';' in message.BCC:
            c = message.BCC.split(";")
        else:
            c = [message.BCC]
    except:
        c = []
    
    d = a + b + c
    
    d = list(filter(None, d))
        
    return d


#%% Name and Address dictionary

def recipient_names_and_addresses(message):
    """ Create dictionary of names and addresses for recipients
    
    Args:
        arg1 message (<class 'win32com.client.CDispatch'>): outlook item
    
    Returns:
        Dictionary of name (key) and address (value) fields
    """
    dict_recipients = {}
    
    # loop over recipients
    try:
        for n in range(message.Recipients.Count):
            
            dict_recipients[message.Recipients.Item(n+1).Name] = message.Recipients.Item(n+1).Address
    except:
        pass
    
    return dict_recipients


#%% Sender Name error handling

def sender_name(message):
    """ Return sender name given message with handling for errors
    
    Args:
        arg1 message (<class 'win32com.client.CDispatch'>): outlook item
    
    Return:
        Sender name text
    """
    
    try:
        sender = message.Sender.Name
    except:
        sender = ''
    
    return sender
    
# Turn into dictionary

#%% Outlook scraping function

def outlook_folder_scrape(folder):
    """ Converts from win32com format to a list of dictionaries
    
    Args:
        arg1 folder (win32com.client.CDispatch): It should be noted that this is specifically a default folder, not Items within
    
    Returns:
        List of Dataframes
    """
    
    # Initialise List
    df_list = []
    
    messages = folder.Items
    
    # Loop through Folder
    for n in range(len(messages)):
        
        if n % 100 == 0:
            print(n)
        
        message_dict = {}
        
        #message_dict['Body'] = message.body
        #message_dict['Subject'] = message.subject
        message_dict['CreationTime'] = convert_to_datetime(messages[n].CreationTime)
        #message_dict['SentTime'] = convert_to_datetime(message.SentOn)
        #message_dict['LastModificationTime'] = convert_to_datetime(message.LastModificationTime)
        message_dict['Recipients'] = recipient_names_and_addresses(messages[n])
        message_dict['Sender'] = sender_name(messages[n])
        #message_dict['CountRecipients'] = messages[n].Recipients.Count
        
        df_list.append(pd.DataFrame(message_dict))
        
    return df_list

df_list = outlook_folder_scrape(Inbox)

df_combined = pd.concat(df_list)

df_combined2.to_excel('mak_inbox.xlsx')






