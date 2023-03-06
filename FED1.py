# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 10:07:55 2020

@author: CWilson
"""

import pandas as pd
import datetime
import matplotlib.pyplot as plt


folder = "C:\\Users\\cwilson\\Documents\\Python\\Fabrication Listing - QC Input"
file = "Fabrication Listing - QC Input 02-23-2021.xlsx"





DROP_BY_UNIT_WT = False
DROP_BY_UNIT_WT_MIN_WT = 100

DROP_BY_QTY = False
DROP_BY_QTY_MAX = 15


LAST_X_DAYS = False
KEEP_LAST_X_DAYS = 30

USE_DATE_RANGE = True
DATE_RANGE = ('2021-01-01', '2021-02-23')



def xldate_to_datetime(xldate):
    temp = datetime.datetime(1899,12,30)
    delta = datetime.timedelta(days=xldate)
    return temp + delta

def keep_last_x_number_of_days(df, x):
    cutoff = max(df["Date"]) - datetime.timedelta(days = x)
    df = df[df["Date"] >= cutoff]
    return df

def keep_date_range(df, date_range):
    start = datetime.datetime.strptime(DATE_RANGE[0], '%Y-%m-%d')
    end = datetime.datetime.strptime(DATE_RANGE[1], '%Y-%m-%d')
    df_start = min(df['Date'])
    df_end = max(df['Date'])

    if df_start >= start:
        start = df_start
    if df_end <= end:
        end = df_end
        
    df = df[df['Date'] >= start]
    df = df[df["Date"] <= end]
    return df

def id_fitters(df):
    fitters = df['Fitter'].unique().tolist()
    new_fitters = []
    for fitter in fitters.copy():
        if "/" in fitter:
            pass
        elif "." in fitter:
            if ".0" in fitter:
                new_fitters.append(fitter)
            else:
                pass
        else:
            new_fitters.append(fitter)
    fitters = new_fitters
    return fitters        

def order_by_sum(series, min_ave_pcs_completed=0, number_of_days=1):
    s = series.sort_values(ascending=False)
    s = s[s/number_of_days > min_ave_pcs_completed]
    return s

def histogram(series, num_bins=10):
    fig, ax = plt.subplots()
    n, bins, patches = ax.hist(series, num_bins)
    ax.set_xlabel("")
    ax.set_ylabel("# of Occurences")
    ax.set_title("")
    plt.show()

def barchart_of_production_by_person(series, title="", ylabel="", note=""):
    x = range(len(series))
    y = series.values
    xlabels = series.index
    fig, ax = plt.subplots()
    ax.bar(x,y, align='center')
    ax.set_ylabel(ylabel)
    ax.set_xlabel("ID #")
    ax.set_title(title)
    ax.set_xticks(x)
    plt.xticks(rotation=90)
    ax.set_xticklabels(xlabels)
    plt.text(int(0.5*max(x)), int(0.8*max(y)), note, fontsize=8)
    plt.tight_layout()
    plt.show()


def x_axis_friendly_date_list(df_date_column):
    dates = df_date_column.tolist()
    for i,date in enumerate(dates):
        dates[i] = date.date().strftime('%Y-%m-%d')
    return dates

def timeline_plot(df, qty_or_wt='qty', title=''):
    if qty_or_wt == 'qty':
        y1 = df['Quantity']
        y2 = df['Cumulative qty']
        y1_label = 'Pcs per day'
        y2_label = 'Cumulative pcs'
    else:
        y1 = df['Weight']
        y2 = df['Cumulative wt']
        y1_label = 'Pounds per day'
        y2_label = 'Cumulative pounds'        
    
    dates = x_axis_friendly_date_list(df['Date'])
    
    fig, ax1 = plt.subplots()
    ax1.set_xlabel('Date')
    ax1.set_title(title)
    ax1.set_ylabel(y1_label)
    ax1.bar(dates, y1)
    ax2 = ax1.twinx()
    ax2.plot(dates, y2, color='black')
    ax2.set_ylabel(y2_label)
    
    ax1.set_xticklabels(dates, rotation=90)
    plt.show()


def apply_filter_yield_fitter_qty(df, filter_type=None, amount=None, note=''):
    if filter_type == None:
        add_to_note = ''
    elif filter_type == 'wt':
        df = df[df['unit wt'] > amount]
        add_to_note = "Min wt = " + str(amount)
    elif filter_type == 'qty':
        df = df[df['Quantity'] < amount]
        add_to_note = "Max qty = " + str(amount)
    else:
        add_to_note = ''
    
    output_note = note + '\n' + add_to_note
    
    fitters = id_fitters(df)
    
    fitter_qty = pd.Series()
    for fitter in fitters:
        fitter_qty[str(fitter)] = df['Quantity'][df['Fitter'] == fitter].sum()
    fitter_qty = order_by_sum(fitter_qty)
    fitter_qty = fitter_qty[fitter_qty > 5]
    
    
    return [df, fitter_qty, output_note]
    # apply filter
    
    #identify the fitters


# this = timeline_dfs['CG']['daily']









plants = {}

xl = pd.ExcelFile(folder + "\\" + file)

for sheet in ['FED QC Form', 'CSF QC Form', 'CSM QC Form']:
    # sheet = 'FED QC Form'
    
    
    
    df = xl.parse(sheet)
    df = df.loc[1:,:].reset_index(drop=True)
    df = df.rename(columns={df.columns[0]:'Date'})
    df = df[df['Date'].notna()]
    
    
    
    
    new_dates = []
    for i in df['Date']:
        if type(i) == float or type(i) == int:
            new_dates.append(xldate_to_datetime(i))
        elif type(i) == str:
            new_dates.append(datetime.datetime.strptime(i, "%m/%d/%Y"))
    df['Date'] = new_dates
    
    
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors = 'coerce')
    df = df[df["Quantity"].notna()]
    df = df[df['Fitter'].notna()]
    df['Fitter'] = df['Fitter'].astype(str)
    if 'Weight (**If REV, show REV Weight Only**)' in df.columns:
        df.rename({df.columns[5]:"Weight"}, inplace=True, axis=1)
    
    df["unit wt"] = df['Weight'] / df["Quantity"]    
    
    
    
    fitters = id_fitters(df)
    # fitters = df['Fitter'].unique().tolist()
    # new_fitters = []
    # for fitter in fitters.copy():
    #     if "/" in fitter:
    #         pass
    #     elif "." in fitter:
    #         if ".0" in fitter:
    #             new_fitters.append(fitter)
    #         else:
    #             pass
    #     else:
    #         new_fitters.append(fitter)
    # fitters = new_fitters
    
    
    
    if LAST_X_DAYS:
        df = keep_last_x_number_of_days(df, KEEP_LAST_X_DAYS)
    if USE_DATE_RANGE:
        df = keep_date_range(df, DATE_RANGE)
        
    
    

    
    
    # note = ''
    # if DROP_BY_UNIT_WT:
    #     df = df[df['unit wt'] > DROP_BY_UNIT_WT_MIN_WT]
    #     note = "Minimum wt. = " + str(DROP_BY_UNIT_WT_MIN_WT)
    # if DROP_BY_QTY:
    #     df = df[df['Quantity'] < DROP_BY_QTY_MAX]
    #     note = "Max qty. = " + str(DROP_BY_QTY_MAX)
    # if DROP_BY_UNIT_WT and DROP_BY_QTY:
    #     note = "Minimum wt. = " + str(DROP_BY_UNIT_WT_MIN_WT) + '\n' + "Max qty. = " + str(DROP_BY_QTY_MAX)    
    
        
    
    
    # fitter_qty = pd.Series()
    # for fitter in fitters:
    #     fitter_qty[str(fitter)] = df['Quantity'][df['Fitter'] == fitter].sum()
    
    # fitter_qty = order_by_sum(fitter_qty)
    
    
    # start = min(df['Date'])
    # end = max(df['Date'])
    
    # num_days = (end - start).days
    
    
    
    
    # histogram(df["Quantity"])
    
    # earliest_date = min(df['Date'])
    # if max(df['Date']) - datetime.timedelta(days = KEEP_LAST_X_DAYS+1) < earliest_date:
    #     title_start = earliest_date
    # else:
    #      title_start = max(df['Date']) - datetime.timedelta(days = KEEP_LAST_X_DAYS+1)
    
    
    # prod_since_title = sheet[:3] + " Fitter production (" + datetime.datetime.strftime(start, '%Y-%m-%d') + " to " + datetime.datetime.strftime(end, '%Y-%m-%d') + ")"
    
    plants[sheet[:3]] = {'base_df': df}
    # plants[sheet[:3]] = {'base_df' : df, 
    #                       'fitter_qty' : fitter_qty,
    #                       'note' : note, 
    #                       'title' : prod_since_title}
    
    

    
xl.close()


for plant in plants:
    base_df = plants[plant]['base_df']
    title = plant + " (" + datetime.datetime.strftime(min(base_df['Date']), '%Y-%m-%d') + ' to ' + datetime.datetime.strftime(max(base_df['Date']), '%Y-%m-%d') + ")"
    
    # Filter with WT
    info1_label = 'wt ' + str(DROP_BY_UNIT_WT_MIN_WT)
    info1 = apply_filter_yield_fitter_qty(base_df, 
                                          filter_type = 'wt', 
                                          amount = DROP_BY_UNIT_WT_MIN_WT, 
                                          note='')
    plants[plant][info1_label] = {'df': info1[0],
                                  'fitter_qty': info1[1],
                                  'note':info1[2]}
    barchart_of_production_by_person(info1[1],
                                     title = title,
                                     ylabel = "# of Pieces",
                                     note = info1[2])    
    # Filter with QTY
    info2_label = 'wqty ' + str(DROP_BY_QTY_MAX)
    info2 = apply_filter_yield_fitter_qty(base_df, filter_type = 'qty',
                                          amount = DROP_BY_QTY_MAX,
                                          note='')
    plants[plant][info2_label] = {'df': info2[0], 
                                  'fitter_qty': info2[1],
                                  'note':info2[2]}
    barchart_of_production_by_person(info2[1],
                                     title = title,
                                     ylabel = "# of Pieces",
                                     note = info2[2])

    # Filter with both QTY & WT    
    info3_label = info1_label + " & " + info2_label
    info3 = apply_filter_yield_fitter_qty(plants[plant]['wt 100']['df'], 
                                         filter_type = 'qty', 
                                         amount = DROP_BY_QTY_MAX, 
                                         note = plants[plant]['wt 100']['note'])
    plants[plant][info3_label] = {'df': info3[0], 
                                  'fitter_qty': info3[1], 
                                  'note':info3[2]}    

    barchart_of_production_by_person(info3[1],
                                     title = title,
                                     ylabel = "# of Pieces",
                                     note = info3[2])
    



""" IM SURE THIS WILL NEED TO FIX AFTER THE CHANGES MADE IN THE FOR LOOP ABOVE"""
"""
timeline_dfs = {}
for plant in plants:
    for fitter in plants[plant]['fitter_qty'].index:
        timeline_df = plants[plant]['base_df']
        timeline_df = timeline_df[timeline_df['Fitter'] == fitter]
        raw_df = timeline_df[['Date', 'Quantity', 'Weight']]
        
        dates = []
        for date in raw_df['Date']:
            if date not in dates:
                dates.append(date)
        daily_df = pd.DataFrame(columns = raw_df.columns)
        for i,date in enumerate(dates):
            daily_df.loc[i,'Date'] = date
            daily_df.loc[i,'Quantity'] = raw_df['Quantity'][raw_df['Date'] == date].sum(axis=0)
            daily_df.loc[i,'Weight'] = raw_df['Weight'][raw_df['Date'] == date].sum(axis=0)
        
        daily_df['Cumulative qty'] = daily_df['Quantity'].cumsum()
        daily_df['Cumulative wt'] = daily_df['Weight'].cumsum()
        timeline_dfs[fitter] = {'raw':raw_df, 'daily':daily_df}
    



timeline_plot(timeline_dfs['2178']['daily'], qty_or_wt='wt', title='2178')

"""



""" Need to create dataframe with every date in it so that there is no skipping around
when adding dates that arnt previously listed """
"""
plt.figure()
for fitter in plants['CSM']['fitter_qty'].index.tolist()[:3]:
    this_df = timeline_dfs[fitter]['daily']
    dates = x_axis_friendly_date_list(this_df['Date'])
    if len(dates) == 1:
        continue
    plt.plot(dates, this_df['Cumulative qty'])
plt.xticks(rotation=90)
plt.show()
"""



# apply the date dropping techinques before the filters
# but do the date dropping after the for loop for reading excel


































