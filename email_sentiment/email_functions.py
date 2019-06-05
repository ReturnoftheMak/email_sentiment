# -*- coding: utf-8 -*-
"""
Created on Thu May 30 14:56:43 2019

@author: Makhan.Gill
"""

#%% Standard Imports

import datetime
import pandas as pd


#%% Datetime conversion

def convert_to_datetime(windows_time):
    """ Convert the pywintypes.datetime to a pandas datetime format.
    Note that instead of just stripping out the timezone it may need to be utilised
    if this is to become usable across timezones.

    Args:
        arg1 windows_time (pywintypes.datetime): Timestamp from outlook item

    Returns:
        datetime
    """

    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            datetime_py = datetime.datetime.strptime(str(windows_time).rstrip("+00:00"), fmt)
            return datetime_py
        except ValueError:
            pass


#%% Recipient dictionary

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


#%% Sender Name

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


#%% Sender Address

def sender_address(message):
    """ Return sender address given message with handling for errors

    Args:
        arg1 message (<class 'win32com.client.CDispatch'>): outlook item

    Return:
        Sender address text
    """

    try:
        if message.Class == 43:
            if message.SenderEmailType == "EX":
                sender = message.Sender.GetExchangeUser().PrimarySmtpAddress
            else:
                sender = message.SenderEmailAddress
        else:
            sender = message.Sender.Address
    except:
        sender = ''

    return sender


#%% Top level folder scrape

def outlook_folder_scrape(folder):
    """ Converts from win32com format to a list of dataframes

    Args:
        arg1 folder (win32com.client.CDispatch): It should be noted that
        this is specifically a default folder, not Items within

    Returns:
        List of Dataframes
    """

    # Initialise List
    df_list = []

    messages = folder.Items

    # Loop through Folder
    for n in range(len(messages)):

        try:
            if n % 100 == 0:
                print(n)

            message_dict = {}

            #message_dict['Body'] = message.body
            #message_dict['Subject'] = message.subject
            message_dict['CreationTime'] = convert_to_datetime(messages[n].CreationTime)
            message_dict['SentTime'] = convert_to_datetime(messages[n].SentOn)
            message_dict['LastModificationTime'] = convert_to_datetime(messages[n].LastModificationTime)
            message_dict['Recipients'] = recipient_names_and_addresses(messages[n])
            message_dict['SenderName'] = sender_name(messages[n])
            message_dict['SenderAddress'] = sender_address(messages[n])
            #message_dict['CountRecipients'] = messages[n].Recipients.Count

            df_list.append(pd.DataFrame(message_dict))

        except:
            pass

    return df_list


#%% Subfolder scrape

def outlook_subfolder_scrape(folder):
    """ Converts from win32com format to a list of dataframes

    Args:
        arg1 folder (win32com.client.CDispatch): It should be noted that
        this is specifically a default folder, not Items within

    Returns:
        List of Dataframes
    """
    
    # Initialise List
    df_list = []
    
    for folder in find_all_subfolders(folder):
        
        try:
            messages = folder.Items
            
            # Loop through subfolder
            for n in range(len(messages)):
        
                try:
                    if n % 100 == 0:
                        print(n)
        
                    message_dict = {}
        
                    #message_dict['Body'] = message.body
                    #message_dict['Subject'] = message.subject
                    message_dict['CreationTime'] = convert_to_datetime(messages[n].CreationTime)
                    message_dict['SentTime'] = convert_to_datetime(messages[n].SentOn)
                    message_dict['LastModificationTime'] = convert_to_datetime(messages[n].LastModificationTime)
                    message_dict['Recipients'] = recipient_names_and_addresses(messages[n])
                    message_dict['SenderName'] = sender_name(messages[n])
                    message_dict['SenderAddress'] = sender_address(messages[n])
                    #message_dict['CountRecipients'] = messages[n].Recipients.Count
        
                    df_list.append(pd.DataFrame(message_dict))
        
                except:
                    pass

                
        except:
            break
    
    return df_list


#%% Iterate through all possible folders

def find_all_subfolders(folder):
    """ Given a folder, find all subfolders
    
    Args:
        arg1 folder (): xxx
    
    Returns:
        list of subfolders?
    """
    
    # Initialise list
    folder_list = []
    
    # Loop through top level folders
    for top_level in folder.Folders:
        
        folder_list.append(top_level)
        
        # Check for next level
        for second_level in top_level.Folders:
            
            folder_list.append(second_level)
            
            # 3rd set of subs
            for third_level in second_level.Folders:
            
                folder_list.append(third_level)
                
                # 4th level
                for fourth_level in third_level.Folders:
                    
                    folder_list.append(fourth_level)
                    
                    # 5th level
                    for fifth_level in fourth_level.Folders:
                    
                        folder_list.append(fifth_level)
    
    return folder_list
