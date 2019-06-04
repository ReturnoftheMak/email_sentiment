# -*- coding: utf-8 -*-
"""
Created on Thu May 30 14:57:40 2019

@author: Makhan.Gill
"""

#%% Imports

import getpass
import pandas as pd
import win32com.client
from email_functions import outlook_folder_scrape, outlook_subfolder_scrape



#%% Initial Setup

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)

sent = outlook.GetDefaultFolder(5)


#%% Run code to scrape

df_list_inbox = outlook_folder_scrape(inbox)

df_list_inbox_subfolder = outlook_subfolder_scrape(inbox)

df_list_sent = outlook_folder_scrape(sent)


#%% Concatenate and format

df_inbox = pd.concat(df_list_inbox, axis=0, join='outer').reset_index().rename(
        columns={'index':'Recipient', 'Recipients':'Address'})

df_inbox_subs = pd.concat(df_list_inbox_subfolder, axis=0, join='outer').reset_index().rename(
        columns={'index':'Recipient', 'Recipients':'Address'})

df_sent = pd.concat(df_list_sent, axis=0, join='outer').reset_index().rename(
        columns={'index':'Recipient', 'Recipients':'Address'})


#%% Export

filepath = r'\\svrtcs04\Syndicate Data\Actuarial\Data\Emails\email_sentiment\email_scrape\\'

df_inbox.to_csv(filepath + getpass.getuser() + '_inbox' + '.csv')

df_inbox_subs.to_csv(filepath + getpass.getuser() + '_inboxsub' + '.csv')

df_sent.to_csv(filepath + getpass.getuser() + '_sent' + '.csv')
