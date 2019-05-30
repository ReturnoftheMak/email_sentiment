# -*- coding: utf-8 -*-
"""
Created on Thu May 30 14:57:40 2019

@author: Makhan.Gill
"""

#%% Imports

import win32com.client
from email_functions import outlook_folder_scrape
import pandas as pd
import getpass


#%% Initial Setup

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

Inbox = outlook.GetDefaultFolder(6)

Sent = outlook.GetDefaultFolder(5)


#%% Run code to scrape

df_list_inbox = outlook_folder_scrape(Inbox)

df_list_sent = outlook_folder_scrape(Sent)


#%% Concatenate and format

df_inbox = pd.concat(df_list_inbox, axis=0, join='outer').reset_index().rename(columns = {'index':'Recipient','Recipients':'Address'})

df_sent = pd.concat(df_list_sent, axis=0, join='outer').reset_index().rename(columns = {'index':'Recipient','Recipients':'Address'})


#%% Export

filepath = r''

df_inbox.to_csv(filepath + getpass.getuser() + '_inbox')

df_sent.to_csv(filepath + getpass.getuser() + '_sent')








