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
from matplotlib.patches import Patch
import textwrap


SAVEFILES = False
color_map = {'MD': 'red', 'DE': 'blue', 'TN': 'green'}
bad_color_map = {'Invalid Employee ID': 'red', 
             'Employee Did Not Work in The Month': 'indigo', 
             'Employee Works at Different Shop': 'grey'}

today = datetime.datetime.now().strftime('%Y-%m-%d')

# df_file = Path(r'C:/Users/Netadmin/documents/FitterWelderPerformanceCSVs/all_both_2025-03_2025-04-18-18-15-57.csv')

# main_df = pd.read_csv(df_file, index_col=0)
# # this is a place holder until I figure out how to alert on when employees are terminated but being credited with work /
# # did not work any during that month but are credited
# main_df = main_df[~main_df['direct'].isna()]



out_folder = Path().home() / 'Documents'/ 'FitterWelderStatsCharts' / today
if not os.path.exists(out_folder):
    out_folder.mkdir(parents=True, exist_ok=True)

def determine_location_and_classification(main_df, state=None, classification=None):
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
    
def weight_by_employee(main_df, state=None, classification=None, topN=25, min_tons=3, SAVEFILES=SAVEFILES):
    col = 'Tonnage' if 'Tonnage' in main_df else 'Tons'
    filename = 'WeightByEmployee_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
    # Sort by Tonnage to match chart order
    df = df.sort_values(col, ascending=False)
    
    title_text = f'{col} Completed by {classification}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[df[col] > min_tons]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(8, 10))
    plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    xlims = plt.gca().get_xlim()
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage + 1, i, loc, va='center')
   # show the tonnage completed if we are passing a state
    else:
        for i, (tonnage) in enumerate(df[col]):
            ha = 'right'
            location = tonnage - 0.1
            if tonnage < xlims[1] * 0.2:
                ha = 'left'
                location = tonnage + 0.2
            plt.text(location, i, f"{tonnage:.1f}", va='center', ha=ha)
    
            
    # Labels and title
    plt.xlabel(col)
    # plt.title(title_text)
            
    
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None
    

# weight_by_employee('MD','Fitter', None)
# weight_by_employee('MD','Welder', None)
# weight_by_employee('DE','Fitter', None)
# weight_by_employee('DE','Welder', None)
# weight_by_employee('TN','Fitter', None)
# weight_by_employee('TN','Welder', None)

# weight_by_employee(classification='Fitter', topN=30)
# weight_by_employee(classification='Welder', topN=30)
#%% Number of Defects

def defects_by_employee(main_df, state=None, classification=None, topN=25, SAVEFILES=SAVEFILES):
    col = 'Defect Quantity'
    filename = 'DefectsByEmployee_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
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
    plt.figure(figsize=(7.5, 10))
    plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage +0.1, i, loc, va='center')
            
    # Labels and title
    plt.xlabel(col)
    # plt.title(title_text)
            
    
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None

# defects_by_employee('MD','Fitter', None)
# defects_by_employee('MD','Welder', None)
# defects_by_employee('DE','Fitter', None)
# defects_by_employee('DE','Welder', None)
# defects_by_employee('TN','Fitter', None)
# defects_by_employee('TN','Welder', None)
# defects_by_employee(classification='Fitter', topN=30)
# defects_by_employee(classification='Welder', topN=30)


def defects_by_employee_both_classifications(main_df, state=None, topN=25, SAVEFILES=SAVEFILES):
    col = 'Defect Quantity'
    classification = 'Combination'
    filename = 'DefectsByEmployeeBothClassifcations_' + file_name_suffix(state, classification) + '.png'
    df_fit = determine_location_and_classification(main_df, state, 'Fitter')
    df_weld = determine_location_and_classification(main_df, state, 'Welder')
    
    df_fit = df_fit[['Name','Location',col]]
    df_weld = df_weld[['Name','Location',col]]

    df = pd.merge(left=df_fit,
                  right=df_weld,
                  left_on=('Name','Location'),
                  right_on=('Name','Location'),
                  how='outer',
                  suffixes=('_Fit','_Weld'))    
    
    df['Total'] = df[[f'{col}_Fit', f'{col}_Weld']].sum(axis=1)
    df = df.sort_values('Total', ascending=False)
    
    # # Sort by Tonnage to match chart order
    # df = df.sort_values(col, ascending=False)
    
    title_text = f'{classification} {col}'
    if not state is None:
        title_text += f' for {state}'
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    df = df[df['Total'] > 0]
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    plt.figure(figsize=(7.5, 10))
    # plt.barh(df['Name'], df[col], color=colors, alpha=0.5)
    
    fit_col  = f"{col}_Fit"
    weld_col = f"{col}_Weld"
        
    plt.barh(
        df['Name'],
        df[fit_col],
        alpha=1.0,
        label='Fit',
        color=colors
    )
    
    # stacked bars
    plt.barh(
        df['Name'],
        df[weld_col],
        left=df[fit_col],
        alpha=0.3,
        label='Weld',
        color=colors,
        # hatch='/'
    )
    
    for i, (fit, weld) in enumerate(zip(df[fit_col], df[weld_col])):
        # Fit label (center of fit segment)
        if fit > 0:
            plt.text(
                fit / 2,
                i,
                f"{fit:.1f}",
                ha='center',
                va='center',
                fontsize=8,
                color='black'
            )
    
        # Weld label (center of weld segment)
        if weld > 0:
            plt.text(
                fit + weld / 2,
                i,
                f"{weld:.1f}",
                ha='center',
                va='center',
                fontsize=8,
                color='black'
            )
    
    colour = colors.iloc[0] if state is not None else 'gray'
    legend_handles = [
    Patch(facecolor=colour, edgecolor='black', label='FIT'),
    Patch(facecolor=colour, edgecolor='black', hatch='', label='WELD', alpha=0.3)
    ]
    
    plt.legend(
        handles=legend_handles,
        loc='upper right',
        frameon=True
    )
    
    plt.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    # add in the state tag to each bar
    if state is None:
        # Adding location labels (optional, to match your plot)
        for i, (tonnage, loc) in enumerate(zip(df[col], df['Location'])):
            plt.text(tonnage +0.1, i, loc, va='center')
            
    # Labels and title
    plt.xlabel(col)
    # plt.title(title_text)
            
    
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None

#%%

def tonnage_per_piece_by_employee(main_df, state=None, classification=None, topN=25, min_qty=None, SAVEFILES=SAVEFILES):
    col = 'Average Tonnage per Piece'
    filename = 'AverageTonsPerPiece_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
    
    low_qty = df[df['Quantity'] < 1]

    # set the quantity to unique quantity if it is less than 1
    df.loc[low_qty.index, 'Quantity'] = df['Unique Quantity']
    # figure out which column to use
    if 'Tonnage' in df.columns:
        df[col] = df['Tonnage'] / df['Quantity']
    elif 'Tons' in df.columns:
        df[col] = df['Tons'] / df['Quantity']
    
    
    df[col] = df[col].replace(np.inf, 0)
    
    # Sort by col to match chart order
    df = df.sort_values(col, ascending=False)
    
    # df = df[df[col] > 0.1]
    
    title_text = f'{classification} {col}'
    if not state is None:
        title_text += f' for {state}'
        
    # cut out any entries with less than min_pieces
    if not min_qty is None:
        df = df[df['Quantity'] >= min_qty]
    
    # get a top remaining entries
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    
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
    # plt.title(title_text)
    
            
    
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None
    

# tonnage_per_piece_by_employee('MD','Fitter', None)
# tonnage_per_piece_by_employee('MD','Welder', None)
# tonnage_per_piece_by_employee('DE','Fitter', None)
# tonnage_per_piece_by_employee('DE','Welder', None)
# tonnage_per_piece_by_employee('TN','Fitter', None)
# tonnage_per_piece_by_employee('TN','Welder', None)
# tonnage_per_piece_by_employee(classification='Fitter', topN=30)
# tonnage_per_piece_by_employee(classification='Welder', topN=30)


#%% Compare Hours worked to Hours earned

def earned_hours_by_employee(main_df, state=None, classification=None, topN=25, min_hours=10, SAVEFILES=SAVEFILES):
    col1 = 'Earned Hours'
    col2 = 'Total Hours'
    col2a = 'Direct Hours'
    filename = 'EarnedHoursAndTotalDirectHours_' + file_name_suffix(state, classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
    df = df.rename(columns={'total':col2, 'direct':col2a})
    
    # must have accomplished atleast 10 hours
    df = df[df['Earned Hours'] > min_hours]

    # Get top N by Total Hours or Earned Hours
    df = df.sort_values(col1, ascending=False)
    if not topN is None:
        df = df.iloc[:topN]

    # df = df[[col1, col2, col2a,, col2b 'Name', 'Location']].copy()
    
    # Color mapping
    colors = df['Location'].map(color_map)

    # Plotting
    fig, (ax, ax2) = plt.subplots(nrows=1, ncols=2, sharey=True, figsize=(8, 10))
    
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
        # ax.text(earned + 5, i, loc, va='center', fontsize=8)
        ax2.text(total + 3, i, f'{int(total)} Hours', va='center', fontsize=8)
        # ax2.text(direct * 0.9, i, f"Direct = {int(direct)}", ha='right', va='center', fontsize=8)


    # Titles and labels
    ax.set_xlabel(col1)
    ax2.set_xlabel(col2)
    ax.set_title(col1)
    ax2.set_title('Total Hours &\nDirect Hours (dark shaded)')
    # plt.title('Earned Hours & Total Hours')

    ax.set_xlim(left=0)
    ax2.set_xlim(left=0)
    # trying to pad a little more space onto the total hours
    current_xlim = ax2.get_xlim()
    ax2.set_xlim((current_xlim[0], current_xlim[1]*1.18))

    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(df['Name'])
    
    

    # Tight layout & save/show
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None

# earned_hours_by_employee('MD','Fitter', None)
# earned_hours_by_employee('MD','Welder', None)
# earned_hours_by_employee('DE','Fitter', None)
# earned_hours_by_employee('DE','Welder', None)
# earned_hours_by_employee('TN','Fitter', None)
# earned_hours_by_employee('TN','Welder', None)
# earned_hours_by_employee(classification='Fitter', topN=30)
# earned_hours_by_employee(classification='Welder', topN=30)

#%% Compare Direct/Indirect/Missed Hours

def hours_comparison_by_employee(main_df, state=None, classification=None, topN=25, SAVEFILES=SAVEFILES):
    col1 = 'Total Hours'
    col2 = 'Direct Hours'
    col3 = 'Indirect Hours'
    col4 = 'Other Hours'
    col5 = 'Missed Hours'
    filename = 'AllHourTypeComparisonByEmployee_' + file_name_suffix(state, classification) + '.png'
    # so that we can pass in a fake {start}-{edn} value for file_name_suffix
    if classification is not None and '-' in classification:
        classification = None
    df = determine_location_and_classification(main_df, state, classification)
    df = df.rename(columns={'total':col1, 'direct':col2, 'indirect':col3, 'missed':col4, 'notcounted':col5})
    df = df[['Name', 'Location', col1, col2, col3, col4, col5]]



    # get xlimits before slicing
    xlims = {
        col1: df[col1].max() * 1.05,
        col2: df[col2].max() * 1.05,
        col3: df[col3].max() * 1.05,
        col4: df[col4].max() * 1.05,
        col5: df[col5].max() * 1.05,
    }
    


    # This is how we order to slice
    df = df.sort_values(col1, ascending=False)
    if topN is None: 
        pass
    elif isinstance(topN, (tuple, list)):
        df = df.iloc[topN[0] : topN[1]]
    else:
        df = df.iloc[:topN]

    
    # This is how we order to display
    df = df.sort_values(col1, ascending=True)

    # Color mapping
    colors = df['Location'].map(color_map)

    # Plotting
    fig, (ax, ax2, ax3, ax4, ax5) = plt.subplots(nrows=1, ncols=5, sharey=True, figsize=(13, 16))
    
    ax.set_xlim(0, xlims[col1])
    ax2.set_xlim(0, xlims[col2])
    ax3.set_xlim(0, xlims[col3])
    ax4.set_xlim(0, xlims[col4])
    ax5.set_xlim(0, xlims[col5])
    
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
        ax.text(location, i, f"{int(total)}", va='center', fontsize=10)
        
        # only add label if value greater than 1
        if direct > 1:
            # put in middle of bar
            location = direct / 2
            # if value is less than 1/5th of the plot width, move it out some 
            if direct < ax2lims[1] * 0.2:
                location = direct + 10
            ax2.text(location, i, f'{int(direct)}', va='center', fontsize=10)
            
        if indirect > 1:
            location = indirect / 2
            if indirect < ax3lims[1] * 0.2:
                location = indirect + 10
            ax3.text(location, i, f'{int(indirect)}', va='center', fontsize=10)
            
        if other > 1:
            location = other / 2
            if other < ax4lims[1] * 0.2:
                location = other + 3
            ax4.text(location, i, f'{int(other)}', va='center', fontsize=10)
            
        if missed > 1:
            location = missed / 2
            if missed < ax5lims[1] * 0.2:
                location = missed + 3
            ax5.text(location, i, f'{int(missed)}', va='center', fontsize=10)            
            

    ax.tick_params(axis='y', labelsize=14) 
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
    # Title being handled in PDF generator
    # fig.suptitle(f"{state if state else 'All Locations'} Hour Breakdown", fontsize=16, y=0.98)
      
    
    # plt.title('Earned Hours & Total Hours')

    ax.set_xlim(left=0)
    ax2.set_xlim(left=0)
    # trying to pad a little more space onto the totla hours

    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(df['Name'])

    # Tight layout & save/show
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename, df
    else:
        plt.show()
        return None
    

# hours_comparison_by_employee('MD',None, None)
# hours_comparison_by_employee('DE',None, None)
# hours_comparison_by_employee('TN',None, None)



#%% plot X=Total Hours, y=Earned Hours

def scatter_total_hours_vs_earned_hours(main_df, state=None, classification=None, topN=25, SAVEFILES=SAVEFILES):
    colx = 'Total Hours'
    coly = 'Earned Hours'
    filename = 'ScatterHoursByEmployee_' + file_name_suffix(state, classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
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
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None
    
# scatter_total_hours_vs_earned_hours('MD','Fitter', None)
# scatter_total_hours_vs_earned_hours('MD','Welder', None)
# scatter_total_hours_vs_earned_hours('DE','Fitter', None)
# scatter_total_hours_vs_earned_hours('DE','Welder', None)
# scatter_total_hours_vs_earned_hours('TN','Fitter', None)
# scatter_total_hours_vs_earned_hours('TN','Welder', None)

#%% Direct Labor Efficiency -->  Earned Hours / Direct Hours

def direct_labor_efficiency(main_df, state=None, classification=None, topN=25, SAVEFILES=SAVEFILES):
    col = 'DL Efficiency'
    col_hours = 'direct'
    filename = 'DirectLaborEfficiency_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
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
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None

# direct_labor_efficiency('MD','Fitter', None)
# direct_labor_efficiency('MD','Welder', None)
# direct_labor_efficiency('DE','Fitter', None)
# direct_labor_efficiency('DE','Welder', None)
# direct_labor_efficiency('TN','Fitter', None)
# direct_labor_efficiency('TN','Welder', None)
# direct_labor_efficiency(classification='Fitter', topN=30)
# direct_labor_efficiency(classification='Welder', topN=30)


#%% DL Efficiency + Total Labor Efficiency -->  Earned Hours / Total Hours

def total_direct_labor_efficiency(main_df, state=None, classification=None, topN=25, min_hours=10, SAVEFILES=SAVEFILES):
    col1 = 'DL Efficiency'
    col2 = 'TTL Efficiency'
    col1_hours = 'direct' if 'direct' in main_df.columns else 'Direct Hours'
    col2_hours = 'total' if 'total' in main_df.columns else 'Total Hours'
    filename = 'DirectAndTotalEfficiency_' + file_name_suffix(state,classification) + '.png'
    df = determine_location_and_classification(main_df, state, classification)
    df[col1] = df['Earned Hours'] / df[col1_hours]
    df[col2] = df['Earned Hours'] / df[col2_hours]
    # must have worked 40 total hours
    df = df[df[col2_hours] > 40]
    # Sort by Tonnage to match chart order
    df = df.sort_values(col1, ascending=False)
    
    # must have wiorked atleast 10 hours
    df = df[df[col1_hours] > min_hours]
    
    
    if not topN is None:
        # get top n
        df = df.iloc[:topN]
        
    # must have both efficiencies above 5%
    df = df[(df[col1] > 0.05) & (df[col2] > 0.05)]
    # get rid of infinitys!
    df = df[(df[col1] != np.inf) & (df[col2] != np.inf)]
    # must have 20 direct hours logged
    # df = df[df[col1_hours] > 20]
    
    
    # Color mapping
    colors = df['Location'].map(color_map)
    
    # Plotting
    fig, (ax, ax2) = plt.subplots(ncols=2, sharey=True, figsize=(8,10))
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
    # # this puts the # of hours they worked on the bars - kinda confusing
    # for i, (dl_eff, ttl_eff, direct_hours, total_hours) in enumerate(zip(df[col1], df[col2], df[col1_hours], df[col2_hours])):
        
    #     location = dl_eff / 2
    #     print_value = f"{int(direct_hours)}"
    #     if dl_eff > ax_xlim[1] * 0.55:
    #         print_value = 'Direct Hours = ' + print_value
    #     elif dl_eff > ax_xlim[1] * 0.35:
    #         print_value = 'Direct = ' + print_value
            
    #     if dl_eff < ax_xlim[1] * 0.25:
    #         location = dl_eff + 0.25
        
    #     ax.text(location, i, print_value, ha='center', va='center', fontsize=8, color='black')
        
    #     location = ttl_eff / 2
    #     print_value = f"{int(total_hours)}"
    #     if ttl_eff > ax2_xlim[1] * 0.55:
    #         print_value = 'Total Hours = ' + print_value
    #     elif ttl_eff > ax2_xlim[1] * 0.35:
    #         print_value = 'Total = ' + print_value
            
    #     if ttl_eff < ax2_xlim[1] * 0.25:
    #         location = ttl_eff + 0.25
        
    #     ax2.text(location, i, print_value, ha='center', va='center', fontsize=8, color='black')
            
        
    ax.set_xlim(ax_xlim[0], ax_xlim[1] * 1.05)
    ax2.set_xlim(ax2_xlim[0], ax2_xlim[1] * 1.05)
    # Labels and title
    ax.set_xlabel(col1)
    ax2.set_xlabel(col2)
    ax.set_title(col1)
    ax2.set_title(col2)
    
    # MAIN TITLE
    # fig.suptitle(f"{state if state else 'All Locations'} {classification} Efficiency", fontsize=16, y=0.98)
    
    # Adjust layout to leave room for suptitle
    # plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None
        

# total_direct_labor_efficiency('MD','Fitter', None)
# total_direct_labor_efficiency('MD','Welder', None)
# total_direct_labor_efficiency('DE','Fitter', None)
# total_direct_labor_efficiency('DE','Welder', None)
# total_direct_labor_efficiency('TN','Fitter', None)
# total_direct_labor_efficiency('TN','Welder', None)
# total_direct_labor_efficiency(classification='Fitter', topN=30)
# total_direct_labor_efficiency(classification='Welder', topN=30)

#%%


def boxPlot_EarnedHours(main_df, state, classificaiton, topN, SAVEFILES=SAVEFILES):
    return None

#%%

def bad_EmployeeCredit(bad_df,  state, classification, SAVEFILES=SAVEFILES):
    filename = 'EmployeesCreditedWithBadData_' + file_name_suffix(state,classification) + '.png'
    
    bad_df = bad_df.copy()
    bad_df['Name'] = bad_df['Name'].fillna('Unknown')
    bad_df['Location'] = bad_df['Location'].fillna(state)
    bad_df_g = bad_df[['Name','Location','reason','Earned Hours','Quantity','Weight']].groupby(['Name','Location','reason']).sum()
    df = bad_df_g.reset_index()
    df['Tons'] = df['Weight'] / 2000
    
    df = df.sort_values(['reason','Earned Hours'], ascending=False)  
    
    # Color mapping
    colors = df['reason'].map(bad_color_map)
    
    
    fig, (ax,ax2,ax3) = plt.subplots(ncols=3, sharey=True, figsize=(8, 5))
    ax.barh(df['Name'], df['Earned Hours'], color=colors, alpha=1)
    ax2.barh(df['Name'], df['Quantity'], color=colors, alpha=1)
    ax3.barh(df['Name'], df['Tons'], color=colors, alpha=1)
    
    
    ax.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax2.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    ax3.grid(axis='x', color='gray', linestyle='--', linewidth=0.5)
    
    
    ax_xlims = ax.get_xlim()
    bars = ax.patches  # Get all the bars from ax
    
    for bar, (loc, reason, eh) in zip(bars, zip(df['Location'], df['reason'], df['Earned Hours'])):
        if loc != state:
            y = bar.get_y() + bar.get_height() / 2
            x = bar.get_width()
            x_offset = ax_xlims[1] * 0.09 if x < ax_xlims[1] * 0.3 else -x / 2
            ax.text(x + x_offset, y, loc, ha='center', va='center', fontsize=8, color='black')
    
            
    # Labels and title
    ax.set_xlabel('Earned Hours')
    ax2.set_xlabel('Quantity')
    ax3.set_xlabel('Tons')
    # plt.title(title_text)
    
    
    ''' calculate totals '''
    earned_hours = df['Earned Hours'].sum()
    quantity = df['Quantity'].sum()
    tons = df['Tons'].sum()
    
    
    # Legend below plot
    legend_elements = [Patch(facecolor=color, label=reason) for reason, color in bad_color_map.items()]
    fig.legend(
        handles=legend_elements,
        title='Reason',
        loc='lower center',
        ncol=3,
        bbox_to_anchor=(0.5, 0.07),
        frameon=True,  # Show the frame
        edgecolor='black',  # Set border color to black
    )
    
    # Shrink layout to make space for legend
    plt.tight_layout(pad=2.0)

    fig.subplots_adjust(bottom=0.3)
    
    # plt.tight_layout()
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return {'filepath':out_folder / filename,
                'earned_hours':earned_hours,
                'quantity':quantity,
                'tons':tons}
    else:
        plt.show()
        return None
    
# bad_EmployeeCredit(bad_df, 'TN','Fitter')
#
#%% 
def bad_entriesPieChart(main_df, bad_df, state, classification, SAVEFILES=SAVEFILES):
    filename = 'BadEntriesPieChart_' + file_name_suffix(state, classification) + '.png'

    bad_df = bad_df.copy()

    bad_df['Name'] = bad_df['Name'].fillna('Unknown')
    bad_df['Location'] = bad_df['Location'].fillna(state)
    bad_df_g = bad_df[['Name', 'Location', 'reason', 'Earned Hours', 'Quantity', 'Weight']].groupby(['Name', 'Location', 'reason']).sum()
    df = bad_df_g.reset_index()
    df['Tons'] = df['Weight'] / 2000

    df = df.sort_values(['reason', 'Earned Hours'], ascending=False)

    # Aggregate totals per reason
    reason_grouped = df[['reason','Earned Hours','Tons','Quantity']].groupby('reason').sum()
    for i in bad_color_map.keys():
        if not i in reason_grouped.index:
            reason_grouped.loc[i,:] = 0
    
    ''' calculations to show in chat '''
    earned_hours = reason_grouped['Earned Hours'].sum()
    quantity = reason_grouped['Quantity'].sum()
    tons = reason_grouped['Tons'].sum()
    
    
    
    df_from_main= determine_location_and_classification(main_df, state, classification)
    
    df_from_main['reason'] = 'Acceptable'
    df_from_main['Tons'] = df_from_main['Weight'] / 2000
    
    df_from_main = df_from_main[['reason','Earned Hours','Quantity','Tons']].groupby('reason').sum()
    
    
    
    reason_grouped = pd.concat([reason_grouped, df_from_main])
    
    # Use 'Earned Hours' for the pie chart
    pie_data = reason_grouped['Earned Hours']
    labels = pie_data.index.tolist()

    # Complete color map including Acceptable
    temp_bad_color_map = bad_color_map.copy()
    temp_bad_color_map['Acceptable'] = 'green'
    colors = [temp_bad_color_map[label] for label in labels]

    def custom_autopct(pct):
        return f'{pct:.1f}%' if pct > 5 else ''

    fig, ax = plt.subplots(1, 1, figsize=(8, 5))

    wedges, texts, autotexts = ax.pie(
        pie_data,
        labels=None,
        colors=colors,
        startangle=140,
        autopct=custom_autopct,
        textprops={'fontsize': 10},
        pctdistance=0.85
    )

    # Move labels outside for slices < 5%
    total = pie_data.sum()
    for i, (wedge, pct) in enumerate(zip(wedges, pie_data)):
        angle = (wedge.theta2 + wedge.theta1) / 2
        if (pct / total) * 100 < 5:
            x = wedge.r * 1.2 * np.cos(np.deg2rad(angle))
            y = wedge.r * 1.2 * np.sin(np.deg2rad(angle))
            label = f'{(pct / total) * 100:.1f}%'
            if (pct / total) > 0:
                ax.text(x, y, label, ha='center', va='center', fontsize=9, color='black')

    plt.tight_layout()
    ax.set_title("Percentage of Earned Hours", fontsize=12)
    
    

    plt.tight_layout()

    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return {'filepath':out_folder / filename,
                'earned_hours':earned_hours,
                'quantity':quantity,
                'tons':tons}
    else:
        plt.show()
        return None
    
# bad_entriesPieChart(main_df, bad_df, state, classification)
# main_df = pdfreport.main_df.copy()
# bad_df = pdfreport.bad_dfs['Welder']

#%%
def bad_entriesPieChartAndTable(main_df, bad_df, state, classification, SAVEFILES=SAVEFILES):
    filename = 'BadEntriesPieChart_' + file_name_suffix(state, classification) + '.png'

    bad_df = bad_df.copy()

    bad_df['Name'] = bad_df['Name'].fillna('Unknown')
    bad_df['Location'] = bad_df['Location'].fillna(state)
    bad_df_g = bad_df[['Name', 'Location', 'reason', 'Earned Hours', 'Quantity', 'Weight']].groupby(['Name', 'Location', 'reason']).sum()
    df = bad_df_g.reset_index()
    df['Tons'] = df['Weight'] / 2000

    df = df.sort_values(['reason', 'Earned Hours'], ascending=False)

    # Aggregate totals per reason (this version will be used for the table)
    reason_grouped = df[['reason','Earned Hours','Tons','Quantity']].groupby('reason').sum()
    for i in bad_color_map.keys():
        if i not in reason_grouped.index:
            reason_grouped.loc[i, :] = 0

    reason_grouped = reason_grouped.sort_index()

    # Save this version for the table
    reason_grouped_for_table = reason_grouped.copy()

    ''' Calculations for summary '''
    earned_hours = reason_grouped['Earned Hours'].sum()
    quantity = reason_grouped['Quantity'].sum()
    tons = reason_grouped['Tons'].sum()

    # Add total row
    total_row = pd.DataFrame({
        'Earned Hours': [earned_hours],
        'Tons': [tons],
        'Quantity': [quantity]
    }, index=['TOTAL'])
    table_df = pd.concat([reason_grouped_for_table, total_row])

    # Prepare pie data with Acceptable data included
    df_from_main = determine_location_and_classification(main_df, state, classification)
    df_from_main['reason'] = 'Acceptable'
    df_from_main['Tons'] = df_from_main['Weight'] / 2000
    df_from_main = df_from_main[['reason', 'Earned Hours', 'Quantity', 'Tons']].groupby('reason').sum()
    reason_grouped_with_acceptable = pd.concat([reason_grouped, df_from_main])

    pie_data = reason_grouped_with_acceptable['Earned Hours']
    labels = pie_data.index.tolist()

    # Complete color map including Acceptable
    temp_bad_color_map = bad_color_map.copy()
    temp_bad_color_map['Acceptable'] = 'green'
    colors = [temp_bad_color_map[label] for label in labels]

    def custom_autopct(pct):
        return f'{pct:.1f}%' if pct > 5 else ''

    # Setup side-by-side figure
    fig, (ax_pie, ax_table) = plt.subplots(1, 2, figsize=(11, 6), gridspec_kw={'width_ratios': [1.2, 1]})

    # Pie chart
    wedges, texts, autotexts = ax_pie.pie(
        pie_data,
        labels=None,
        colors=colors,
        startangle=140,
        autopct=custom_autopct,
        textprops={'fontsize': 14},
        pctdistance=0.85
    )
    
   
    # Outside labels for small slices
    total = pie_data.sum()
    for i, (wedge, pct, reason) in enumerate(zip(wedges, pie_data, pie_data.index)):
        angle = (wedge.theta2 + wedge.theta1) / 2
        if (pct / total) * 100 < 5:
            x = wedge.r * 1.2 * np.cos(np.deg2rad(angle))
            y = wedge.r * 1.2 * np.sin(np.deg2rad(angle))
            label = f'{(pct / total) * 100:.1f}%'
            if (pct / total) * 100 > 0.5:
                ax_pie.text(x, y, label, ha='center', va='center', fontsize=14, color='black')
        
                
        

    ax_pie.set_title("Amount Of\nEarned Hours", fontsize=16)

    # Table on the right
    ax_table.axis('off')  # Hide axes
    
    # Wrap reason labels
    wrapped_labels = [textwrap.fill(str(idx), width=11) for idx in table_df.index]
    
    # Format table data
    table_data = [[wrapped] + [f'{v:.1f}' for v in row] for wrapped, (_, row) in zip(wrapped_labels, table_df.iterrows())]
    
    column_labels = ['Reason', 'Earned Hours', 'Tons', 'Quantity']
    table = ax_table.table(cellText=table_data, colLabels=column_labels, loc='center', cellLoc='center')
    
    # Adjust styling
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1.1, 5)  # Wider rows and more vertical space per row
    
    # Optional: make header bold
    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(weight='bold')
            
    # Light grey background for the last row (TOTAL)
    num_rows = len(table_df) + 1  # +1 for header
    for col in range(len(column_labels)):
        table[num_rows - 1, col].set_facecolor('#c0c0c0')

    plt.tight_layout()

    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return {'filepath': out_folder / filename,
                'earned_hours': earned_hours,
                'quantity': quantity,
                'tons': tons}
    else:
        plt.show()
        return None

# bad_entriesPieChartAndTable(main_df, bad_df, state, classification)
# main_df = pdfreport.main_df
# bad_df = pdfreport.bad_dfs['Welder]

#%%


def mom_hours_worked_by_shop(dict_of_dfs, state=None, hours_type='Total Hours', classification=None, month_0_name='March', month_0_year=2025, SAVEFILES=SAVEFILES):
    if classification is None:
        classification_text = ''
    else:
        classification_text = f"_Class-{classification}"
    
    filename = f'MonthOverMonthLine_State-{state}{classification_text}_{hours_type.replace(' ','')}.png'
    
    if isinstance(month_0_year, str):
        month_0_year = int(month_0_year)
        
    month_start = datetime.datetime.strptime(month_0_name,'%B')
    month_start = datetime.datetime(month_0_year, month_start.month, 1)
    
    
    x_display_names = []
    
    dfs = dict_of_dfs.copy()
    dfs = dict(sorted(dfs.items()))
    for k in dfs:
        df = dfs[k].copy()
        df = df[df['Location'] == state]
        # make sure we are gpoing to get the right class when doing the drop_duplicates
        if classification is not None:
            # use a custom sort key so that the value of Classification=classification is always the first option if it is available
            df = df.sort_values(
                by=['Name', 'Classification'],
                key=lambda col: col.map(lambda x: 0 if x == classification else 1) if col.name == 'Classification' else col
            )
            
        df = df.drop_duplicates(subset='Name', keep='first')
            
        # only get people who actually contributed to fabrication via number of earned hours 
        df = df[(df['Earned Hours'] > 0) | (df['Weight'] > 0) | (df['Quantity'] > 0)]
        df = df[['Name',hours_type,'Classification']]
        df= df.set_index('Name')
        dfs[k] = df
        
        
        months_ago = month_start
        for _ in range(k):
            months_ago = months_ago - datetime.timedelta(days=1)
            months_ago = datetime.datetime(months_ago.year, months_ago.month, 1)
    
        x_display_names.append(months_ago.strftime('%b %Y'))
            
        
        
    keys = list(dfs.keys())
    keys.sort()
    # get the first dataframe as the current month
    df = dfs[keys[0]].copy()
    # add suffix _0
    df = df.add_suffix('_0')
    # for dataframes 1:end
    for i in range(1,len(keys)):
        # get the joining df
        right_df = dfs[keys[i]].copy()
        # add its suffix
        right_df = right_df.add_suffix(f'_{i}')
        # do a left join
        df = pd.merge(left=df, right=right_df,
                      left_index=True, right_index=True,
                      how='left')
        
    if not classification is None:
        df = df[df['Classification_0'] == classification]
        
    df = df.drop(columns = [i for i in df.columns if 'Classification' in i])
        
    # only get employees who worked the zero month
    df = df[df.iloc[:,0] > 0]
        
    
    if hours_type == 'Missed Hours':
        # for this type of hours we need to include zeros
        avg_per_column = df.mean(axis=0)
        mean_nonzero = df.stack().mean()

    else:    
        df = df.fillna(0)
        avg_per_column = df[df > 0].mean(axis=0)
        mean_nonzero = df[df != 0].stack().mean()
    

    
    avg_series = pd.Series(avg_per_column)
    avg_series = avg_series[::-1]  # Reverse so Total Hours_6 is on the left
    
        
    fig, ax = plt.subplots(1, 1, figsize=(8, 5))
    
    # Plot line
    ax.plot(avg_series.index, avg_series.values, marker='o', label=f'Average {hours_type}')
    
    # Horizontal mean line
    ax.axhline(y=mean_nonzero, color='red', linestyle='--', label=f'Average: {mean_nonzero:.0f}')
    
    # Set custom x-tick labels
    ax.set_xticks(avg_series.index)
    ax.set_xticklabels(x_display_names[::-1], rotation=45)
    
    # Labels and title
    # ax.set_xlabel('Month')
    ax.set_ylabel('Average Hours per Employee')
    # ax.set_title(f'Month-Over-Month Comparison of {hours_type}')
    
    # Grid, legend, layout
    ax.legend()
    ax.grid(True)
    fig.tight_layout()
    
    if SAVEFILES:
        plt.savefig(out_folder / filename, dpi=300)
        plt.close()
        return out_folder / filename
    else:
        plt.show()
        return None