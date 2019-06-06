# -*- coding: utf-8 -*-
"""
Created on Fri May 31 15:52:50 2019

@author: Makhan.Gill
"""

#%% Library Import

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt


#%% Set Variables

InpathChris = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_inbox.csv'
InpathMatt = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Matthew.Rout_inbox.csv'
InpathMak = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Makhan.Gill_inbox.csv'
InpathCarly = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\carly.perry_inbox.csv'
InpathGrace = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Grace.Owens_inbox.csv'

InSpathChris = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_inboxsub.csv'
InSpathMatt = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Matthew.Rout_inboxsub.csv'
InSpathMak = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Makhan.Gill_inboxsub.csv'
InSpathCarly = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\carly.perry_inboxsub.csv'
InSpathGrace = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Grace.Owens_inboxsub.csv'

SentpathChris = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_sent.csv'
SentpathMatt = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Matthew.Rout_sent.csv'
SentpathMak = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Makhan.Gill_sent.csv'
SentpathCarly = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\carly.perry_sent.csv'
SentpathGrace = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Grace.Owens_sent.csv'

#df_chris = pd.read_csv(InpathChris)
#df_chris['User'] = 'Chris.Gallagher'
#
#df_matt = pd.read_csv(InpathMatt)
#df_matt['User'] = 'Matthew.Rout'
#
#df_mak = pd.read_csv(InpathMak)
#df_mak['User'] = 'Makhan.Gill'

Inboxes = [InpathChris,
           InpathMatt,
           InpathMak,
           InpathCarly,
           InpathGrace,
           InSpathChris,
           InSpathMak,
           InSpathCarly]

df_list = [pd.read_csv(inbox) for inbox in Inboxes]

Combined_df = pd.concat(df_list)

Combined_df['count'] = 1

df_grouped = Combined_df.groupby(['Recipient', 'SenderName'], as_index=False).sum()

df_grouped = df_grouped[['Recipient', 'SenderName', 'count']]

df_all = df_grouped[df_grouped['Recipient'].isin(['Gallagher, Christopher', 'Gill, Makhan', 'Rout, Matthew', 'Owens, Grace', 'Perry, Carly']) & (df_grouped['count'] >= 40)]


df_dirtyNames = df_all[df_all['SenderName'].str.contains("-")]
df_dirtyNames2 = df_dirtyNames['SenderName'].str.split('-', 1, expand=True)
df_dirtyNames2.columns = ['DirtyPart', 'CleanName']
df_allMerge = pd.merge(df_all, df_dirtyNames2, left_index=True, right_index=True, how='left')

df_allMerge['CleanName'] = df_allMerge['CleanName'].fillna(value=df_allMerge['SenderName'])

df_all['SenderName'] = df_allMerge['CleanName']

#Fudge
df_all['SenderName'] = df_all['SenderName'].str[-20:]


#%% Draw network

A = nx.Graph()

A.add_nodes_from(list(df_all['SenderName'].unique()))

tuples = [tuple(x) for x in df_all.values]

A.add_weighted_edges_from(tuples)


#%% Draw spring diagram

fig = plt.figure(figsize=(30, 30))
ax = plt.subplot(111)
limits = plt.axis('off')
ax.set_title('Dependencies', fontsize=18)

nx.draw_spring(A, arrows=False, with_labels=True, alpha=0.7, width=0.5, font_size=15)


#%% Save plot

plt.tight_layout()
plt.savefig(r"\\svrtcs04\Syndicate Data\Actuarial\Data\Emails\email_sentiment\email_sentiment\Graph4.png", format="PNG")
plt.show()
