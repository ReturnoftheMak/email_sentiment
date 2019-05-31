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

path = r'M:\Actuarial\Data\Emails\email_sentiment\email_scrape\Chris.Gallagher_inbox.csv'

df = pd.read_csv(path)

df['count'] = 1

df_grouped = df.groupby(['Recipient', 'SenderName'], as_index=False).sum()

df_grouped = df_grouped[['Recipient', 'SenderName', 'count']]

df_chris = df_grouped[(df_grouped['Recipient'] == 'Gallagher, Christopher') & (df_grouped['count'] >= 50)]


#%% Draw network

A = nx.Graph()

A.add_nodes_from(list(df_chris['SenderName'].unique()))

tuples = [tuple(x) for x in df_chris.values]

A.add_weighted_edges_from(tuples)


#%%

fig = plt.figure(figsize=(200,200))
ax = plt.subplot(111)
limits=plt.axis('off')
ax.set_title('Dependencies', fontsize=18)

nx.draw_spring(A, arrows=False, with_labels=True, alpha=0.6, width=0.15, font_size=7)


#%%

plt.tight_layout()
plt.savefig("Graph.png", format="PNG")
plt.show()



