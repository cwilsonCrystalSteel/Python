# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 17:33:26 2025

@author: Netadmin
"""

import matplotlib.pyplot as plt
import pandas as pd
from pathlib import Path
import datetime
import os

today = datetime.datetime.now().strftime('%Y-%m-%d')

df_file = Path(r'C:\Users\Netadmin\Documents\FitterWelderPerformanceCSVs\all_both_02-02-2025_to_02-28-2025_2025-04-15-17-32-33.csv')

main_df = pd.read_csv(df_file, index_col=0)

color_map = {'MD': 'red', 'DE': 'blue', 'TN': 'green'}

out_folder = Path().home() / 'Documents'/ 'FitterWelderStatsCharts' / today
if not os.path.exists(out_folder):
    out_folder.mkdir(parents=True, exist_ok=True)

def determine_location_and_classification(state=None, classification=None):
    df = main_df.copy()
    if state == None:
        pass
    else:
        df = df[df['Location'] == state]
                
    if classification == None:
        pass
    else:
        df= df[df['Classification'] == classification]
                
    return df

def file_name_suffix(state=None, classification=None):
    if state is None:
        state = 'All'
    if classification is None:
        classification = 'Both'
        
    return f"State-{state}_{classification}"
    
    
    
def weight_by_employee(state=None, classification=None, topN=25):
    filename = 'WeightByEmployee_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    # Sort by Tonnage to match chart order
    df = df.sort_values('Tonnage', ascending=False)
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 10))
    bars = plt.barh(df['Name'], df['Tonnage'], color=colors)
    
    # Labels and title
    plt.xlabel('Tonnage')
    plt.title(f'Weight Completed by {classification}')
    
    # # Adding location labels (optional, to match your plot)
    # for i, (tonnage, loc) in enumerate(zip(df['Tonnage'], df['Location'])):
    #     plt.text(tonnage + 1, i, loc, va='center')
    
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()


weight_by_employee('MD','Fitter', None)
weight_by_employee('MD','Welder', None)

weight_by_employee(classification='Fitter', topN=40)
weight_by_employee(classification='Welder', topN=40)
