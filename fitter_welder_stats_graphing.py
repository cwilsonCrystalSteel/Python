# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 17:33:26 2025

@author: Netadmin
"""

import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
import pandas as pd
from pathlib import Path
import datetime
import os
import numpy as np

today = datetime.datetime.now().strftime('%Y-%m-%d')

df_file = Path(r'C:/Users/Netadmin/documents/FitterWelderPerformanceCSVs/all_both_2025-03_2025-04-18-18-15-57.csv')

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
                
    if classification is not None:
        df = df[df['Classification'] == classification]
    else:
        # Get the first occurrence of each employee
        df = df.sort_values('Date' if 'Date' in df.columns else df.columns[0])  # optional: sort by date if available
        df = df.drop_duplicates(subset='Name', keep='first')
                
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
    plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
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

weight_by_employee(classification='Fitter', topN=30)
weight_by_employee(classification='Welder', topN=30)
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
    plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
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
defects_by_employee(classification='Fitter', topN=30)
defects_by_employee(classification='Welder', topN=30)

#%%

def tonnage_per_piece_by_employee(state=None, classification=None, topN=25):
    col = 'Tonnage per Piece'
    filename = 'AverageTonsPerPiece_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    
    low_qty = df[df['Quantity'] < 1]
    df.loc[low_qty.index, col] = df['Tonnage'] / df['Unique Quantity']
    df.loc[low_qty.index, 'Quantity'] = df['Unique Quantity']
    
    # Sort by col to match chart order
    df = df.sort_values(col, ascending=False)
    
    # df = df[df[col] > 0.1]
    
    title_text = f'{classification} {col}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 6))
    plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage +0.1, i, loc, va='center')
    
    xlims = plt.gca().get_xlim()
    
    # all to try and line up the location of PCs.
    for i, (value, qty) in enumerate(zip(df[col], df['Quantity'])):
        formatted = f"{int(qty) if qty.is_integer() else round(qty, 1)}"            
        location = value / 2
        if value > xlims[1] * 0.17:
            formatted += ' Pcs.'
        if value < xlims[1] * 0.07:
            location = value + 0.15
        
        plt.text(location, i, f'{formatted}',
             ha='center', va='center', fontsize=8, color='black')  
            
    # Labels and title
    plt.xlabel(col)
    plt.title(title_text)
    
            
    
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()
    

tonnage_per_piece_by_employee('MD','Fitter', None)
tonnage_per_piece_by_employee('MD','Welder', None)
tonnage_per_piece_by_employee('DE','Fitter', None)
tonnage_per_piece_by_employee('DE','Welder', None)
tonnage_per_piece_by_employee('TN','Fitter', None)
tonnage_per_piece_by_employee('TN','Welder', None)
tonnage_per_piece_by_employee(classification='Fitter', topN=30)
tonnage_per_piece_by_employee(classification='Welder', topN=30)


#%% Compare Hours worked to Hours earned

def earned_hours_by_employee(state=None, classification=None, topN=25):
    col1 = 'Earned Hours'
    col2 = 'Total Hours'
    col2a = 'Direct Hours'
    filename = 'EarnedHoursAndTotalDirectHours' + file_name_suffix(state, classification) + '.png'
    df = determine_location_and_classification(state, classification)
    df = df.rename(columns={'total':col2, 'direct':col2a})

    # Get top N by Total Hours or Earned Hours
    df = df.sort_values(col1, ascending=False)
    if not topN is None:
        df = df.iloc[:topN]

    # df = df[[col1, col2, col2a,, col2b 'Name', 'Location']].copy()
    
    # Color mapping
    colors = df['Location'].map(color_map)

    # Plotting
    fig, (ax, ax2) = plt.subplots(nrows=1, ncols=2, sharey=True, figsize=(12, 10))
    
    # Plot Earned Hours (left side)
    ax.barh(df['Name'], df[col1], color=colors, label=col1, alpha=0.5)

    # Plot Total Hours (right side)
    ax2.barh(df['Name'], df[col2], color=colors, alpha=0.5, label=col2)
    ax2.barh(df['Name'], df[col2a], color='black', alpha=0.25, label=col2a)
    
    # Grid only on total hours axis
    ax.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax2.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)

    # Add text labels on both bars
    for i, (earned, total, direct, loc) in enumerate(zip(df[col1], df[col2], df[col2a], df['Location'])):
        ax.text(earned + 5, i, loc, va='center', fontsize=8)
        ax2.text(total + 3, i, f'{int(total)} Hours', va='center', fontsize=8)
        # ax2.text(direct * 0.9, i, f"Direct = {int(direct)}", ha='right', va='center', fontsize=8)


    # Titles and labels
    ax.set_xlabel(col1)
    ax2.set_xlabel(col2)
    ax.set_title(col1)
    ax2.set_title('Total Hours & Direct Hours (dark shaded)')
    # plt.title('Earned Hours & Total Hours')

    ax.set_xlim(left=0)
    ax2.set_xlim(left=0)
    # trying to pad a little more space onto the totla hours
    current_xlim = ax2.get_xlim()
    ax2.set_xlim((current_xlim[0], current_xlim[1]*1.11))

    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(df['Name'])

    # Tight layout & save/show
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()
    

earned_hours_by_employee('MD','Fitter', None)
earned_hours_by_employee('MD','Welder', None)
earned_hours_by_employee('DE','Fitter', None)
earned_hours_by_employee('DE','Welder', None)
earned_hours_by_employee('TN','Fitter', None)
earned_hours_by_employee('TN','Welder', None)
earned_hours_by_employee(classification='Fitter', topN=30)
earned_hours_by_employee(classification='Welder', topN=30)

#%% Compare Direct/Indirect/Missed Hours

def hours_comparison_by_employee(state=None, classification=None, topN=25):
    col1 = 'Total Hours'
    col2 = 'Direct Hours'
    col3 = 'Indirect Hours'
    col4 = 'Other Hours'
    col5 = 'Missed Hours'
    filename = 'AllHourTypeComparisonByEmployee_' + file_name_suffix(state, classification) + '.png'
    df = determine_location_and_classification(state, classification)
    df = df.rename(columns={'total':col1, 'direct':col2, 'indirect':col3, 'missed':col4, 'notcounted':col5})

    # Get top N by Total Hours or Earned Hours
    df = df.sort_values(col1, ascending=False)
    if not topN is None:
        df = df.iloc[:topN]

    
    # Color mapping
    colors = df['Location'].map(color_map)

    # Plotting
    fig, (ax, ax2, ax3, ax4, ax5) = plt.subplots(nrows=1, ncols=5, sharey=True, figsize=(12, 10))
    
    # Plot Earned Hours (left side)
    ax.barh(df['Name'], df[col1], color=colors, label=col1, alpha=0.5)
    ax2.barh(df['Name'], df[col2], color=colors, label=col2, alpha=0.5)
    ax3.barh(df['Name'], df[col3], color=colors, label=col3, alpha=0.5)
    ax4.barh(df['Name'], df[col4], color=colors, label=col4, alpha=0.5)
    ax5.barh(df['Name'], df[col5], color=colors, label=col5, alpha=0.5)

    # Grid only on total hours axis
    ax.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax2.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax3.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax4.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax5.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    
    
    
    axlims = ax.get_xlim()
    ax2lims = ax2.get_xlim()
    ax3lims = ax3.get_xlim()
    ax4lims = ax4.get_xlim()
    ax5lims = ax5.get_xlim()

    # Add text labels on bars
    for i, (total, direct, indirect, other, missed) in enumerate(zip(df[col1], df[col2], df[col3], df[col4], df[col5])):
        
        location = total / 2
        if total < axlims[1] * 0.2:
            location = total + 10
        ax.text(location, i, f"{int(total)}", va='center', fontsize=8)
        
        # only add label if value greater than 1
        if direct > 1:
            # put in middle of bar
            location = direct / 2
            # if value is less than 1/5th of the plot width, move it out some 
            if direct < ax2lims[1] * 0.2:
                location = direct + 10
            ax2.text(location, i, f'{int(direct)}', va='center', fontsize=8)
            
        if indirect > 1:
            location = indirect / 2
            if indirect < ax3lims[1] * 0.2:
                location = indirect + 10
            ax3.text(location, i, f'{int(indirect)}', va='center', fontsize=8)
            
        if other > 1:
            location = other / 2
            if other < ax4lims[1] * 0.2:
                location = other + 3
            ax4.text(location, i, f'{int(other)}', va='center', fontsize=8)
            
        if missed > 1:
            location = missed / 2
            if missed < ax5lims[1] * 0.2:
                location = missed + 3
            ax5.text(location, i, f'{int(missed)}', va='center', fontsize=8)            
            


    # Titles and labels
    ax.set_xlabel(col1)
    ax2.set_xlabel(col2)
    ax3.set_xlabel(col3)
    ax4.set_xlabel(col4)
    ax5.set_xlabel(col5)
    
    ax.set_title(col1)
    ax2.set_title(col2)
    ax3.set_title(col3)
    ax4.set_title(col4)
    ax5.set_title(col5)
    
    # MAIN TITLE
    fig.suptitle(f"{state if state else 'All Locations'} Hour Breakdown", fontsize=16, y=0.98)
      
    
    # plt.title('Earned Hours & Total Hours')

    ax.set_xlim(left=0)
    ax2.set_xlim(left=0)
    # trying to pad a little more space onto the totla hours

    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(df['Name'])

    # Tight layout & save/show
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()
    

hours_comparison_by_employee('MD',None, None)
hours_comparison_by_employee('DE',None, None)
hours_comparison_by_employee('TN',None, None)


#%% plot X=Total Hours, y=Earned Hours

def scatter_total_hours_vs_earned_hours(state=None, classification=None, topN=25):
    colx = 'Total Hours'
    coly = 'Earned Hours'
    filename = 'ScatterHoursByEmployee_' + file_name_suffix(state, classification) + '.png'
    df = determine_location_and_classification(state, classification)
    df = df.rename(columns={'total': colx})

    # Get top N by Total Hours or Earned Hours
    df = df.sort_values(coly, ascending=False)
    
    if not topN is None:
        df = df.iloc[:topN]

    df = df[[colx, coly, 'Name', 'Location']].copy()

    # Sort by y-axis value (Earned Hours) for consistent numbering
    df = df.sort_values(coly, ascending=False).reset_index(drop=True)
    
    df['Point #'] = df.index + 1 

    # Color mapping
    colors = df['Location'].map(color_map)

    # Plotting
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.scatter(df[colx], df[coly], color=colors, alpha=0.5)

    # Annotate each point with its number
    for i, row in df.iterrows():
        ax.text(row[colx]+1.5, row[coly]+10, str(row['Point #']), va='center', fontsize=8)

    # Build the legend
    legend_entries = [f"{int(row['Point #'])}: {row['Name']}" for _, row in df.iterrows()]
    legend_text = "Ordered by Earned Hours:\n\n"
    legend_text += "\n".join(legend_entries)

    # Add legend to side as a text box
    plt.gcf().text(0.8, 0.5, legend_text, va='center', fontsize=8, family='monospace')

    # Grid, labels, title
    ax.grid(axis='both', color='gray', linestyle='--', linewidth=0.5)
    ax.set_xlabel(colx)
    ax.set_ylabel(coly)
    ax.set_title(f'[{state}] [{classification}] Total Hours Worked vs. Earned Hours')
    
    # try to get a little more room on the x-axis for number
    current_xlim = ax.get_xlim()
    ax.set_xlim((current_xlim[0], current_xlim[1]*1.03))

    # Layout & show/save
    plt.tight_layout(rect=[0, 0, 0.8, 1])  # Make room for legend
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()

    
scatter_total_hours_vs_earned_hours('MD','Fitter', None)
scatter_total_hours_vs_earned_hours('MD','Welder', None)
scatter_total_hours_vs_earned_hours('DE','Fitter', None)
scatter_total_hours_vs_earned_hours('DE','Welder', None)
scatter_total_hours_vs_earned_hours('TN','Fitter', None)
scatter_total_hours_vs_earned_hours('TN','Welder', None)

#%% Direct Labor Efficiency -->  Earned Hours / Direct Hours

def direct_labor_efficiency(state=None, classification=None, topN=25):
    col = 'DL Efficiency'
    col_hours = 'direct'
    filename = 'DirectLaborEfficiency_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    df[col] = df['Earned Hours'] / df[col_hours]
    df = df[df[col_hours] > 40]
    # Sort by Tonnage to match chart order
    df = df.sort_values(col, ascending=False)
    
    title_text = f'{classification} {col}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[df[col] > 0.2]
    df = df[df[col] != np.inf]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 6))
    plt.gca().xaxis.set_major_formatter(PercentFormatter(xmax=1))

    plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (x, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(x + 0.1, i, loc, va='center')
    
    
    xlims = plt.gca().get_xlim()
    for i, (efficiency, direct_hours) in enumerate(zip(df[col], df[col_hours])):
        print_value = f"{int(direct_hours)}"
        if efficiency > xlims[1] * 0.3:
            print_value = 'Direct Hours = ' + print_value
        elif efficiency > xlims[1] * 0.15:
            print_value = 'Direct = ' + print_value
        
        plt.text(efficiency / 2, i, print_value,
             ha='center', va='center', fontsize=8, color='black')
            
    plt.gca().set_xlim(xlims[0], xlims[1] * 1.05)
    # Labels and title
    plt.xlabel(col)
    plt.title(title_text)
            
    
    plt.tight_layout()
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()

direct_labor_efficiency('MD','Fitter', None)
direct_labor_efficiency('MD','Welder', None)
direct_labor_efficiency('DE','Fitter', None)
direct_labor_efficiency('DE','Welder', None)
direct_labor_efficiency('TN','Fitter', None)
direct_labor_efficiency('TN','Welder', None)
direct_labor_efficiency(classification='Fitter', topN=30)
direct_labor_efficiency(classification='Welder', topN=30)


#%% DL Efficiency + Total Labor Efficiency -->  Earned Hours / Total Hours

def total_direct_labor_efficiency(state=None, classification=None, topN=25):
    col1 = 'DL Efficiency'
    col2 = 'TTL Efficiency'
    col1_hours = 'direct'
    col2_hours = 'total'
    filename = 'DirectAndTotalEfficiency_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(state, classification)
    df[col1] = df['Earned Hours'] / df[col1_hours]
    df[col2] = df['Earned Hours'] / df[col2_hours]
    df = df[df[col2_hours] > 40]
    # Sort by Tonnage to match chart order
    df = df.sort_values(col1, ascending=False)
    
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[(df[col1] > 0.2) & (df[col2] > 0.2)]
    df = df[(df[col1] != np.inf) & (df[col2] != np.inf)]
    df = df[df[col1_hours] > 20]
    
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    fig, (ax, ax2) = plt.subplots(ncols=2, sharey=True, figsize=(10,8))
    ax.xaxis.set_major_formatter(PercentFormatter(xmax=1))
    ax2.xaxis.set_major_formatter(PercentFormatter(xmax=1))

    ax.barh(df['Name'], df[col1], color=colors, alpha=0.5)
    ax2.barh(df['Name'], df[col2], color=colors, alpha=0.5)
    
    ax.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax2.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (dl_eff, ttl_eff, loc) in enumerate(zip(df[col1], df[col2], df['Location'])):
            ax.text(dl_eff + 0.1, i, loc, va='center')
            ax2.text(ttl_eff + 0.1, i, loc, va='center')
    
    
    ax_xlim = ax.get_xlim()
    ax2_xlim = ax2.get_xlim()
    for i, (dl_eff, ttl_eff, direct_hours, total_hours) in enumerate(zip(df[col1], df[col2], df[col1_hours], df[col2_hours])):
        print_value = f"{int(direct_hours)}"
        if dl_eff > ax_xlim[1] * 0.4:
            print_value = 'Direct Hours = ' + print_value
        elif dl_eff > ax_xlim[1] * 0.2:
            print_value = 'Direct = ' + print_value
        
        ax.text(dl_eff / 2, i, print_value, ha='center', va='center', fontsize=8, color='black')
        
        print_value = f"{int(total_hours)}"
        if ttl_eff > ax2_xlim[1] * 0.4:
            print_value = 'Total Hours = ' + print_value
        elif ttl_eff > ax2_xlim[1] * 0.2:
            print_value = 'Total = ' + print_value
        
        ax2.text(ttl_eff / 2, i, print_value, ha='center', va='center', fontsize=8, color='black')
            
        
    ax.set_xlim(ax_xlim[0], ax_xlim[1] * 1.05)
    ax2.set_xlim(ax2_xlim[0], ax2_xlim[1] * 1.05)
    # Labels and title
    ax.set_xlabel(col1)
    ax2.set_xlabel(col2)
    ax.set_title(col1)
    ax2.set_title(col2)
    
    # MAIN TITLE
    fig.suptitle(f"{state if state else 'All Locations'} {classification} Efficiency", fontsize=16, y=0.98)
    
    # Adjust layout to leave room for suptitle
    # plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.tight_layout()
    
    # plt.savefig(out_folder / filename, dpi=300)
    # plt.close()
    plt.show()

total_direct_labor_efficiency('MD','Fitter', None)
total_direct_labor_efficiency('MD','Welder', None)
total_direct_labor_efficiency('DE','Fitter', None)
total_direct_labor_efficiency('DE','Welder', None)
total_direct_labor_efficiency('TN','Fitter', None)
total_direct_labor_efficiency('TN','Welder', None)
total_direct_labor_efficiency(classification='Fitter', topN=30)
total_direct_labor_efficiency(classification='Welder', topN=30)

