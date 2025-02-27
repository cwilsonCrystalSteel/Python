# -*- coding: utf-8 -*-
"""
Created on Fri Nov 13 07:37:23 2020

@author: CWilson
"""

"""
Purpose:
    Generate plot/graph to show how many pieces each fitter and welder have
    completedsince the start of 2019.
    
Objectives:
    
    1) Make the import fab list column splitter into a function with more capability
        - check if welder present or not
        - count the weight, qty, type of piece (based on pc-mark initial letters)
    
    2) Figure out what kind of pieces to drop and protocol in loop to not count embeds
        -> Starts with "EA" (embed angle) or "EP" (embed plate)
        
    3) Do fitter performance starting at 1/1/2019
        Make it so that the script can run and handle either just fitters or f&w
        
    5) Create a pcs/day metric
    
    4) Create way to track qty, hpp, and pcs/day by month for each guy
    


'''

'''    
FIGURE OUT WHY WHEN YOU RUN ONLY NOVEMBER, IT ONLY HAS 4 WELDERS and 5 FITTERS
    - Because of the argument 'drop_amt' in function drop_cols
        * Need to figure out how to do 'drop_amt' based on number of days worked
'''

'''         
            
NOTES
Realized that the person who welded pieces was not being tracked until April 9, 2020

"""


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from os import listdir
from os.path import isfile, join
import sys
sys.path.append("C:\\Users\\cwilson\\Documents\\Python")

from read_daily_fab_listing_function import compile_daily_fab


# Does not count the pieces produced if the qty exceeds QTY_OVERRIDE
QTY_OVERRIDE_ON = False
QTY_OVERRIDE = 20


daily_fab_python_folder = "C:\\Users\\cwilson\\Documents\\Daily Fabrication\\Daily Fab Python"
files = [f for f in listdir(daily_fab_python_folder) 
            if isfile(join(daily_fab_python_folder,f))]

def pull_employee_hours():
    # let eh = employee_hours
    eh_folder = daily_fab_python_folder + "\\Employee Hours"
    eh_files = [f for f in listdir(eh_folder) if isfile(join(eh_folder,f))]
    eh_files = [f for f in eh_files if f[-5:] == '.xlsx']
    eh_dates = [d[15:-5] for d in eh_files]
    eh_dates = [datetime.strptime(d,"%Y-%m-%d-%H-%M-%S") for d in eh_dates]
    # Pulls the most recent file based on the timestamp on the file name
    eh_file = eh_files[np.argmax(eh_dates)]
    eh_df = pd.read_excel(eh_folder + "\\" + eh_file, header=3)
    return eh_df
eh_df = pull_employee_hours()


def find_employee_sections():
    eh_id = pull_employee_hours()["Week"].dropna()
    for i in eh_id.index:
        if type(eh_id[i]) != str:
            eh_id = eh_id.drop(index=i)
        elif eh_id[i] == 'Period Totals':
            eh_id = eh_id.drop(index=i)
    eh_id = eh_id.to_frame(name='Identifier')
    eh_id = eh_id.reset_index()
    eh_id = eh_id.rename({'index':'eh_df index'}, axis='columns')

    # Analyze identifier columns to get employee name and ID
    for i in eh_id.index:
        a = eh_id["Identifier"][i]
        number_pos = a.find('Number')
        name = a[:number_pos].strip()
        id_num = a[number_pos+8:number_pos+12]
        eh_id.at[i, 'Name'] = name
        eh_id.at[i, 'ID'] = int(id_num)
    return eh_id
eh_id = find_employee_sections()



def pull_employee_codes_names():
    # Let ecn = employee_codes_names
    # View the .txt file in the ecn folder on how to pull the newest ecn csv
    ecn_folder = daily_fab_python_folder + "\\Employee codes & names"
    ecn_files = [f for f in listdir(ecn_folder) if isfile(join(ecn_folder,f))]
    ecn_files = [f for f in ecn_files if f[-4:] == ".csv"]
    ecn_dates = [d[28:-4] for d in ecn_files]
    ecn_dates = [datetime.strptime(d,"%Y-%m-%d-%H-%M-%S") for d in ecn_dates]
    ecn_file = ecn_files[np.argmax(ecn_dates)]
    ecn_df = pd.read_csv(ecn_folder + "\\" + ecn_file, names=["ID","First","Last"], index_col=0)
    ecn_df["Full Name"] = ecn_df["First"] + " " + ecn_df["Last"]
    return ecn_df

ecn_df = pull_employee_codes_names()


fit_tl = pd.DataFrame(columns=['Date'])
weld_tl = pd.DataFrame(columns=['Date'])
tl_idx_tracker = 0 


# File 16 is may 2020
for file in files[-2:]:
    xl = pd.ExcelFile(daily_fab_python_folder + "\\" + file)
    year = file[:4]
    month = file[5:7]
    sheets = xl.sheet_names
    
    for day in sheets:
        # This skips the sheet in the excel file if sheet name is not the number of the day 
        try:
            int(day)
        except:
            continue
        
        df = xl.parse(str(day))
        
        if len(day) == 1:
            day = "0" + day
            
        try:
            date = datetime.strptime(year+'-'+month+'-'+day,'%Y-%m-%d')
        except:
            continue
    
        fit_tl = fit_tl.append({'Date':date}, ignore_index=True)
        weld_tl = weld_tl.append({'Date':date}, ignore_index=True)
        
        # cols = df.columns
        indexer = df.loc[df[df.columns[0]] == "Job #"].index[0]
        new_cols = df.loc[indexer].fillna("split here").tolist()
        split_here_index = new_cols.index("split here")
        df1_cols = new_cols[:split_here_index]
        df1_cols[df1_cols.index("Fitter")+1] = "QCF"
        df1_cols[df1_cols.index("Welder")+1] = "QCW"
        df2_cols = new_cols[split_here_index+1:]
        df2_cols[df2_cols.index("Fitter")+1] = "QCF"
        df2_cols[df2_cols.index("Welder")+1] = "QCW"
        df2_cols = [str(x)+" 1" for x in df2_cols]
        new_cols = df1_cols + ["split here"] + df2_cols
        
        df = df.loc[indexer+1:].reset_index()
        df = df.drop([df.columns[0]], axis=1)
        df.columns = new_cols
        df2_cols = [i for i in df2_cols if i[:5] != "split"]    
        df1 = df[df1_cols]
        df2 = df[df2_cols]
        df2.columns = df1.columns
        todays_dfs = [df1, df2]
        
        weld_idx = weld_tl[weld_tl["Date"]==date].index[0]
        fit_idx = fit_tl[fit_tl["Date"]==date].index[0]
    
        for this_df in todays_dfs:
            for idx in this_df.index:
                qty = this_df.loc[idx, "Qty"]
                if type(qty) != int:
                    continue
                else:
                    fitter = this_df.loc[idx, "Fitter"]
                    welder = this_df.loc[idx, "Welder"]
                    
                    if QTY_OVERRIDE_ON == True:
                        if qty > QTY_OVERRIDE:
                            qty = 0
                    
                    if pd.isnull(fitter) == False:
                        if fitter in fit_tl.columns:
                            if pd.isnull(fit_tl.at[fit_idx, fitter]) == True:
                                fit_tl.at[fit_idx, fitter] =  + qty
                            else:
                                fit_tl.at[fit_idx, fitter] = fit_tl.at[fit_idx, fitter] + qty
                        else:
                            fit_tl.at[weld_idx, fitter] = qty  
                                                   
                    if pd.isnull(welder) == False:
                        if welder in weld_tl.columns:
                            if pd.isnull(weld_tl.at[weld_idx, welder]) == True:
                                weld_tl.at[weld_idx, welder] =  + qty
                            else:
                                weld_tl.at[weld_idx, welder] = weld_tl.at[weld_idx, welder] + qty
                        else:
                            weld_tl.at[weld_idx, welder] = qty       




fcheck = [i for i in fit_tl.columns if (type(i)==int) and (len(str(i))==4)
          or (type(i)==str) and (len(str(i))==2)]
wcheck = [i for i in weld_tl.columns if (type(i)==int) and (len(str(i))==4)
          or (type(i)==str) and (len(str(i))==2)]



def drop_cols(check_list, df, pcs_per_day=(2/3)):
    
    drop_amt = df.index[-1] * pcs_per_day
    
    df1 = df["Date"].to_frame()
    for col in df.columns.tolist():
        if col in check_list:
            if df.sum(axis=0)[col] > drop_amt:
                df1[col] = df[col]
    
    Zak_remove_list = [2189, 2176, 2001, 2084, 2106, 2177, 2015, 2190, 2175]
    for i in Zak_remove_list:
        if i in df1.columns:
            del df1[i]
            
    return df1

fit_tl1 = drop_cols(fcheck, fit_tl)
weld_tl1 = drop_cols(wcheck, weld_tl)


all_fitters_welders = []
for f in fit_tl1.columns[1:]:
    if f not in all_fitters_welders:
        all_fitters_welders.append(f)
        
for w in weld_tl1.columns[1:]:
    if w not in all_fitters_welders:
        all_fitters_welders.append(w)



def order_by_sum(dataframe):
    s = dataframe.sum(axis=0).sort_values(ascending=False)
    cols = ['Date']
    cols.extend(s.index)
    # for col in s.index:
    #     cols.append(col)
    dataframe = dataframe.reindex(cols, axis=1)
    return dataframe

fit_tl2 = order_by_sum(fit_tl1)
weld_tl2 = order_by_sum(weld_tl1)


def ids_to_names(dataframe):
    d = dataframe.copy()
    l = d.columns.to_list()
    for i in l:
        if i in ecn_df.index:
            d = d.rename(columns={i:ecn_df.at[i,"Full Name"]})
            
    return d

fit_tl3 = ids_to_names(fit_tl2)
weld_tl3 = ids_to_names(weld_tl2)







days = [weld_tl3["Date"].min()]
for d in range((weld_tl3["Date"].max() - weld_tl3["Date"].min()).days):
    days.append(weld_tl3["Date"].min() + timedelta(days=(d+1)))



def pull_hours_per_day():
    """
    SOMETHING IS NOT RIGHT FOR NIGHT SHIFT -> HOURS DISPLAYED IN THE EXCEL FILE ARE 
    STRANGE AND I DON'T WANT TO TRY AND FIX IT CURRENTLY (as of 11/15/2020)
        -> Example: Col:2193, index:122:124 (2020-08-31 to 2020-09-02)
    """
    
    # Created the dataframe that has everyones ID # with all of the dates
    # hpd = Hours_Per_Day
    hpd = fit_tl1["Date"].to_frame()
    for guy in all_fitters_welders:
        hpd[guy] = 0
        if type(guy) == str:
            continue
    
        start_idx = eh_id.at[eh_id.index[eh_id["ID"] == guy][0], 'eh_df index']
        for idx in eh_df.index[start_idx:]:
            if eh_df.loc[idx]['Week'] == 'Period Totals':
                break
            elif pd.isnull(eh_df.loc[idx]['In']) == True:
                continue
            else:
                md = eh_df.loc[idx]['In']
                m = md[:md.find('/')]
                d = md[md.find('/')+1:]
                if len(m) == 1:
                    m = '0' + m
                if len(d) == 1:
                    d = '0' + d
                date = datetime.strptime('2020' + m + d, '%Y%m%d')
                while pd.isnull(eh_df.loc[idx]["Day Total"]) == True:
                    idx += 1
                hours = eh_df.loc[idx]["Day Total"]
                hpd_idx = hpd.index[hpd['Date'] == date]
                hpd.at[hpd_idx, guy] = hours
    return hpd

hpd = pull_hours_per_day()          
hpd1 = ids_to_names(hpd)           



def hours_per_piece(fit_or_weld_pcs_completed_df, names=True):
    hpd_df = hpd1
    if names == False:
        hpd_df = hpd
    cols_to_use = fit_or_weld_pcs_completed_df.columns.to_list()
    hpp = hpd_df[cols_to_use].sum() / fit_or_weld_pcs_completed_df.sum()
    return hpp

hpp_fit_names = hours_per_piece(fit_tl3)
hpp_weld_names = hours_per_piece(weld_tl3)



def get_timeline_str(df):
    start = min(df['Date'])
    end = max(df['Date'])
    
    timeline_str = str(start.month) + '-' + str(start.day) + ' to ' + str(end.month) + '-' + str(end.day)
    return timeline_str



def my_plot_bar_chart(df_or_series, title, ylabel=''):
    # Creates a bar chart to display # of pieces completed by employee
    # dict1 = input dict with {"name/ID":Pcs Completed}
    # title = title of plot such that: Title + "From may 1 to october 31"
    
    plot_title = title 
    
    if isinstance(df_or_series, pd.DataFrame):
        s = df_or_series.sum(axis=0)
    else:
        s = df_or_series
    x = range(len(s))
    y = s.values
    xlabels = [i[:12] for i in s.index]
    fig, ax = plt.subplots()
    ax.bar(x,y, align="center")
    ax.set_ylabel(ylabel)
    ax.set_title(plot_title)
    ax.set_xticks(x)
    plt.xticks(rotation=90)
    ax.set_xticklabels(xlabels)
    plt.tight_layout()
    plt.savefig(fname=plot_title, dpi=300, papertype='letter')
    plt.show()








my_plot_bar_chart(fit_tl3, title="Fitters " + get_timeline_str(fit_tl3) , ylabel="# of Pieces")
my_plot_bar_chart(weld_tl3, title="Welders " + get_timeline_str(fit_tl3), ylabel="# of Pieces")

my_plot_bar_chart(hpp_fit_names, title="Fitters Hours Per Piece", ylabel="Hours per Piece")
my_plot_bar_chart(hpp_weld_names, title="Welder Hours Per Piece", ylabel="Hours per Piece")
















