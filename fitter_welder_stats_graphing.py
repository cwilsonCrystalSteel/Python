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

df_file = Path(r'C:/Users/Netadmin/documents/FitterWelderPerformanceCSVs/all_both_2025-03_2025-04-18-12-53-39.csv')

main_df = pd.read_csv(df_file, index_col=0)
# this is a place holder until I figure out how to alert on when employees are terminated but being credited with work /
# did not work any during that month but are credited
main_df = main_df[~main_df['direct'].isna()]


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
    
#%% Tons completed
    
def weight_by_employee(state=None, classification=None, topN=25):
    col = 'Tonnage'
    filename = 'WeightByEmployee_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    # Sort by Tonnage to match chart order
    df = df.sort_values(col, ascending=False)
    
    title_text = f'{col} Completed by {classification}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[df[col] > 3]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 10))
    plt.barh(df['Name'], df[col], color=colors)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage + 1, i, loc, va='center')
            
    # Labels and title
    plt.xlabel(col)
    plt.title(title_text)
            
    
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()
    

weight_by_employee('MD','Fitter', None)
weight_by_employee('MD','Welder', None)
weight_by_employee('DE','Fitter', None)
weight_by_employee('DE','Welder', None)
weight_by_employee('TN','Fitter', None)
weight_by_employee('TN','Welder', None)

weight_by_employee(classification='Fitter', topN=40)
weight_by_employee(classification='Welder', topN=40)
#%% Number of Defects

def defects_by_employee(state=None, classification=None, topN=25):
    col = 'Defect Quantity'
    filename = 'DefectsByEmployee_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    # Sort by Tonnage to match chart order
    df = df.sort_values(col, ascending=False)
    
    title_text = f'{classification} {col}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[df[col] > 0]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 6))
    plt.barh(df['Name'], df[col], color=colors)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage +0.1, i, loc, va='center')
            
    # Labels and title
    plt.xlabel(col)
    plt.title(title_text)
            
    
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()

defects_by_employee('MD','Fitter', None)
defects_by_employee('MD','Welder', None)
defects_by_employee('DE','Fitter', None)
defects_by_employee('DE','Welder', None)
defects_by_employee('TN','Fitter', None)
defects_by_employee('TN','Welder', None)
defects_by_employee(classification='Fitter', topN=40)
defects_by_employee(classification='Welder', topN=40)

#%%

def weight_by_piece_by_employee(state=None, classification=None, topN=25):
    col = 'Tonnage per Piece'
    filename = 'AverageTonsPerPiece_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    # Sort by Tonnage to match chart order
    df = df.sort_values(col, ascending=False)
    
    title_text = f'{classification} {col}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[df[col] > 0]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 6))
    plt.barh(df['Name'], df[col], color=colors)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage +0.1, i, loc, va='center')
            
    # Labels and title
    plt.xlabel(col)
    plt.title(title_text)
            
    
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()

weight_by_piece_by_employee('MD','Fitter', None)
weight_by_piece_by_employee('MD','Welder', None)
weight_by_piece_by_employee('DE','Fitter', None)
weight_by_piece_by_employee('DE','Welder', None)
weight_by_piece_by_employee('TN','Fitter', None)
weight_by_piece_by_employee('TN','Welder', None)
weight_by_piece_by_employee(classification='Fitter', topN=10)
weight_by_piece_by_employee(classification='Welder', topN=12)
