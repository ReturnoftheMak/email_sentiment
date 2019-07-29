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

# Should probably change this to glob and regex at some point

InpathChris = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_inbox.csv'
InpathMatt = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Matthew.Rout_inbox.csv'
InpathMak = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Makhan.Gill_inbox.csv'
InpathCarly = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\carly.perry_inbox.csv'
InpathGrace = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Grace.Owens_inbox.csv'
InpathDC = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\DC_inbox.csv'
InpathDM = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\David.Marshall_inbox.csv'
InpathMB = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\michael.boyce_inbox.csv'


InSpathChris = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_inboxsub.csv'
InSpathMatt = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Matthew.Rout_inboxsub.csv'
InSpathMak = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Makhan.Gill_inboxsub.csv'
InSpathCarly = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\carly.perry_inboxsub.csv'
InSpathGrace = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Grace.Owens_inboxsub.csv'
InSpathDC = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\DC_inbox_subs.csv'
InSpathMB = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\michael.boyce_inboxsub.csv'


SentpathChris = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_sent.csv'
SentpathMatt = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Matthew.Rout_sent.csv'
SentpathMak = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Makhan.Gill_sent.csv'
SentpathCarly = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\carly.perry_sent.csv'
SentpathGrace = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Grace.Owens_sent.csv'
SentpathDC = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\DC_sent.csv'
SentpathDM = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\David.Marshall_sent.csv'
SentpathMB = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\michael.boyce_sent.csv'


#%% Data Manip

# Add to these as necessary, or use a list comprehension with glob if changed above

Inboxes = [InpathChris,
           InpathMatt,
           InpathMak,
           InpathCarly,
           InpathGrace,
           InpathDC,
           InpathDM,
           InpathMB,
           InSpathChris,
           InSpathMatt,
           InSpathMak,
           InSpathCarly,
           InSpathGrace,
           InSpathDC,
           InSpathMB]

Sentboxes = [SentpathChris,
             SentpathMatt,
             SentpathMak,
             SentpathCarly,
             SentpathGrace,
             SentpathDC,
             SentpathDM,
             SentpathMB]


inbox_list = [pd.read_csv(inbox) for inbox in Inboxes]

sent_list = [pd.read_csv(sent) for sent in Sentboxes]

combined_inboxes = pd.concat(inbox_list)

combined_sentboxes = pd.concat(sent_list)

combined_inboxes['count'] = 1

combined_sentboxes['count'] = 1

combined_inboxes['box'] = 'inbox'

combined_sentboxes['box'] = 'sent'

combined_inboxes = combined_inboxes.dropna(axis=0, how='any', subset=['SenderAddress'])

combined_sentboxes = combined_sentboxes.dropna(axis=0, how='any', subset=['SenderAddress'])

combined_emails = pd.concat([combined_inboxes, combined_sentboxes])

combined_emails.to_csv('combined.csv')


#%% Groupby and forming a sent/received dataframe

#combined_inboxes_ch = combined_inboxes[combined_inboxes['SenderAddress'].str.contains('channel2015')]

df_grouped_inbox = combined_inboxes.groupby(['Recipient', 'SenderName'], as_index=False).sum()

df_grouped_sent = combined_sentboxes.groupby(['Recipient', 'SenderName'], as_index=False).sum()

df_grouped_inbox = df_grouped_inbox[['Recipient', 'SenderName', 'count']] #.rename(columns={'count':'Received'})

df_grouped_sent = df_grouped_sent[['Recipient', 'SenderName', 'count']] #.rename(columns={'count':'Sent'})



df_merged = df_grouped_inbox.append(df_grouped_sent)

# Probably some more work neeeded here to make the merge work

act_only = df_merged[(df_merged['Recipient'].isin(['Gallagher, Christopher', 'Gill, Makhan', 'Rout, Matthew', 'Owens, Grace', 'Perry, Carly', 'Marshall, David', 'Charlton, David', 'Boyce, Michael']))
 & (df_merged['SenderName'].isin(['Gallagher, Christopher', 'Gill, Makhan', 'Rout, Matthew', 'Owens, Grace', 'Perry, Carly', 'Marshall, David', 'Charlton, David', 'Boyce, Michael']))]



#%% Create IDs

ids = pd.DataFrame(df_merged['Recipient'].unique())

ids = ids.reset_index()

ids['index'] += 1

ids = ids.rename(columns={'index':'Id', 0:'Label'})


#%% Create Edges

edges_start = df_merged[['Recipient', 'SenderName', 'count']].rename(columns={'SenderName':'source_', 'Recipient':'target_', 'count':'Weight'})

edges_start['Type'] = 'Directed'

edges = edges_start.merge(ids, how='left', left_on='target_', right_on='Label')

edges = edges.merge(ids, how='left', left_on='source_', right_on='Label')

edges = edges.rename(columns={'Id_x':'Target', 'Id_y':'Source'})

edges_output = edges[['Source', 'Target', 'Type', 'Weight']]


#%% Create Nodes

users = r'\\svrtcs04\Syndicate Data\Actuarial\Data\Emails\email_sentiment\email_sentiment\Users.xlsx'

user_data = pd.read_excel(users)

department_lookup = user_data[['PreferredName', 'Department']]

nodes = ids.merge(department_lookup, how='left', left_on='Label', right_on='PreferredName')

nodes_output = nodes[['Id', 'Label', 'Department']]


#%% Output to semicolon delimited files without indices

edges_output.to_csv('edges.csv',sep=';',index=False)

nodes_output.to_csv('nodes.csv',sep=';',index=False)


#%% Test files for Gephi tutorial

pd.read_csv('http://www.martingrandjean.ch/wp-content/uploads/2015/10/Nodes1.csv', sep=';').to_csv('nodes_tutorial.csv',sep=';',index=False)

pd.read_csv('http://www.martingrandjean.ch/wp-content/uploads/2015/10/Edges1.csv', sep=';').to_csv('edges_tutorial.csv',sep=';',index=False)


#%%

mgandmr = df_merged[(df_merged['Recipient'].isin(['Gill, Makhan', 'Rout, Matthew'])) & (df_merged['SenderName'].isin(['Gill, Makhan', 'Rout, Matthew']))]


df_all = df_merged[df_merged['Recipient'].isin(['Gallagher, Christopher', 'Gill, Makhan', 'Rout, Matthew', 'Owens, Grace', 'Perry, Carly', 'Marshall, David', 'Charlton, David', 'Boyce, Michael']) & (df_merged['Received'] >= 25)]


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
