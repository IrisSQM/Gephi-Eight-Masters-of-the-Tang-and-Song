#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Data Cleaning Codes
"""
'''
The data for the social network of the eight masters of the Tang and Song 
are extracted from China Biographical Database Project (CBDB).

This code is intended to prepare two excel tables for visualization in Gephi.

input: query results.xlsx

output: the_eight_gephi.xlsx
'''
import pandas as pd
import os
 
path = '<Your directory path>'

raw = pd.read_excel(path + os.sep + 'query results.xlsx')

nodes = pd.DataFrame()
linked = pd.DataFrame() # Dataset is pairwise so need to record source and target separately
edge = pd.DataFrame()

# Select data to nodes
# Add index person to nodes
nodes[['Id','Label','姓名','Index Year','Dynasty','Index Place']] = raw[['PersonID','Name','姓名','Index Year','Dynasty','Index Place']]

# Add linked person to nodes
linked[['Id','Label','姓名','Index Year','Dynasty','Index Place']] = raw[['NodeID','Linked to','社會關係人姓名',"Node's Index Year",'Node Dynasty','Node Index Place']]

# Merge two df
nodes = pd.concat([nodes,linked])

# Remove duplicates
nodes = nodes.drop_duplicates()

# Set Id as index
nodes = nodes.set_index('Id')

# Mark the eight masters of the Tang and Song
Eight = ['Su Shi','Liu Zongyuan','Han Yu','Ouyang Xiu','Su Xun','Su Zhe','Wang Anshi','Zeng Gong']
nodes['The Eight'] = nodes['Label'].apply(lambda x: x in Eight)

# Select data to edge
edge[['Source','Target','Weight','Edge Dist.','Distance 距離']] = raw[['PersonID','NodeID','Count','Edge Dist.','Distance 距離']]

# Set Source as index
edge = edge.set_index('Source')

# Export tables
writer = pd.ExcelWriter(path + os.sep + 'the_eight_gephi.xlsx', engine = 'xlsxwriter')
nodes.to_excel(writer, sheet_name = 'nodes_table')
edge.to_excel(writer, sheet_name = 'edges_table')

writer.save()