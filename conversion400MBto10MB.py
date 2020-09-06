# -*- coding: utf-8 -*-
""" Conversion from 400 MB to 10 MB data volume
Use:
    1. Save excel data h1 to h8 under C:\Users\xx\Documents\ROH_01042018bis02092020\
    2. Define the folder name "ROH_01042018bis02092020" under the variable fol.
    3. Run Skript.
    4. CSV is given out on desctop.
"""
#%% Import packages
import pandas as pd
import os
import numpy as np

#%% Prepare data

fol = "ROH_01042018bis02092020"

# Save the title of every xlsx file in the folder in a list li
source = "C:\\Users\\xx\\Documents\\{}".format(fol)
li = []
for file in os.listdir(source):
    if file.endswith(".xlsx"):
        # save title of xlsx files in li
        li.append(os.path.join(source, file))
        # Control mechanism: Print all xlsx files in source
        #print(os.source.join(source, file))

#%% Z_HH
"""Since the data is a times series and every doc is formatted the same way, 
the following code takes the datetime column from h1.xlsx and saves it in zeit."""
zeit = pd.read_excel("C:\\Users\\xx\\Documents\\{}\\h1.xlsx".format(fol), 
                     sheet_name='Werte', usecols='K', skiprows=5, names=['Zeit'])
df_li = []
# Extract data from h1 to h8 and save it in a list df_li
for path in li[0:8]:
    # nimm den aktuellen index aus li
    idx = li.index(path)
    # zähle +1 zum index hinzu
    num = idx + 1
    data = pd.read_excel(path, sheet_name='Werte', usecols='O:S', skiprows=5,
                         names=[f'Z_BAD_H{num}@Haus{num}', f'Z_RLT_H{num}@Haus{num}',
                                f'Z_TWW_H{num}@Haus{num}', f'Z_EV_H{num}@Haus{num}', 
                                f'Z_HH{num}@Haus{num}'])
    df_li.append(data)   
# Summarize data of h1 to h8 in a single DataFrame
z_hh = pd.concat(df_li, axis=1)

#%% Concat the datetimecolumn and the values of h1 to h8
frames = [zeit, z_hh]
frames = pd.concat(frames, axis=1)
# Set zeit as an index
result = frames.set_index('Zeit')
# Create csv
result.to_csv(r'C:\Users\xx\Desktop\CSV_01042018bis02092020\' + fol[-19:] + '.csv') 
