# -*- coding: utf-8 -*-
"""
Created on Mon Jun  7 15:14:02 2021

@author: CWilson
"""
#%% Load modules and some functions
from Read_Group_hours_HTML import output_each_clock_entry_job_and_costcode
import pandas as pd
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Fitter_Welder_Stats_functions import get_earned_hours_per_ton
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import datetime
import glob
import os


def do_day_stuff(series_name, df, weekdays):
    day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    if isinstance(weekdays, int):
        this_day = np.array(df[df['Day of Week'] == weekdays][series_name])
        days = day_of_week[weekdays]
    elif isinstance(weekdays, list):
        this_day = np.array(df[df['Day of Week'].isin(weekdays)][series_name])
        days = []
        for day in weekdays:
            days.append(day_of_week[day][:3])
        days = ', '.join(days)
    return {'series':this_day, 'days':days}

def remove_outliers_individually(series, percentile):
    z_score_for_percentile = stats.norm.ppf(percentile)
    series = series[np.abs(stats.zscore(series)) < z_score_for_percentile]
    return series 


def get_amount_of_OT_on_saturdays():
    remove_jobs = ['650-PTO (Shop)','648-Unpaid time off request','500-Office', '651 - Holiday']
    folder = 'c://users/cwilson/downloads/xyz/'
    files = glob.glob(folder+'*.html')
    file = files[0]
    dates = [datetime.datetime.strptime(os.path.basename(f)[:-5],'%m-%d-%Y') for f in files]
    date_dict = {}
    for i,file in enumerate(files):
        date_dict[dates[i]] = {}
        df = pd.read_html(file)[0]
        na_rows = df.index[df[0].isna()]
        group_rows = na_rows - 1
        df = df[~df[0].isna()]
        df = df[~df.index.isin(group_rows)]
        df = df.iloc[1:]
        
        df = df[~df[0].isin(remove_jobs)]
        
        
        totals = df[df.columns[1:]].astype(float).sum()

        date_dict[dates[i]]['Total'] = totals[1]
        date_dict[dates[i]]['Regular'] = totals[2]
        date_dict[dates[i]]['OT'] = totals[3]
    
    df = pd.DataFrame.from_dict(date_dict, orient='index')
    df['Normalized OT'] = df['Regular'] + 1.5 * df['OT']
    return df




#%% Gather data from the sources

hours = output_each_clock_entry_job_and_costcode('c://users/cwilson/downloads/jan1tojune5hours.html')
hours = hours[hours['Job'] != '651 - Holiday']
hours = hours[hours['Job'] != '650 - PTO (Shop)']
hours = hours[hours['Job'] != '648 - Unpaid time off request']
hours = hours[hours['Job'] != '500-Office']


fablisting = grab_google_sheet('CSM QC Form','01/01/2021','06/05/2021')


sold_hours = get_earned_hours_per_ton()['TN']


ot = get_amount_of_OT_on_saturdays()



#%% Clean 1

hours['Date'] = hours['Start'].dt.date

# cut out Robert Burrel, Robert Richardson, Rodger Grant


hours_by_date = hours.groupby(['Date']).sum()['Hours']
hours_by_date = hours_by_date.reset_index()
hours_by_date['Date'] = pd.to_datetime(hours_by_date['Date'])
hours_by_date['Day of Week'] = hours_by_date['Date'].dt.weekday



# NEED TO GET THE FABLISTING DATA AND THEN CONVERT WEIGHT TO EARNED HOURS FOR EACH DAY
fablisting['Date'] = fablisting['Timestamp'].dt.date
fablisting = fablisting.set_index('Job #')
sold_hours = sold_hours.set_index('Job')
fablisting = fablisting.join(sold_hours)

fablisting['Earned Hours'] = (fablisting['Weight'] / 2000) * fablisting['Hrs/Ton']
fablisting_by_date = fablisting.groupby(['Date']).sum()

#%% Calculate 1



df = hours_by_date.set_index('Date').copy()
df = df.join(fablisting_by_date[['Weight','Quantity','Earned Hours']])
df['Tons'] = df['Weight'] / 2000
df['TTL Hours per Ton'] = df['Hours'] / (df['Weight'] / 2000)
df['TTL Efficiency'] = df['Earned Hours'] / df['Hours']


df_ot = df.copy()
k = ot['Normalized OT']
for i in k.index:
    df_ot.loc[i,'Hours'] = k[i]
del i
df_ot['TTL Hours per Ton'] = df_ot['Hours'] / (df_ot['Weight'] / 2000)
df_ot['TTL Efficiency'] = df_ot['Earned Hours'] / df_ot['Hours']


# chop out sundays
df = df[df['Day of Week'] != 6]
df_ot = df_ot[df_ot['Day of Week'] != 6]

drop_days_without_production = False
if drop_days_without_production:
    df = df[~df['Weight'].isna()]
    
df = df.fillna(0)
df_ot = df_ot.fillna(0)

df_by_day = df.groupby(['Day of Week']).sum()
df_by_day['TTL Efficiency'] = df_by_day['Earned Hours'] / df_by_day['Hours']

df_ot_by_day = df_ot.groupby(['Day of Week']).sum()
df_ot_by_day['TTL Efficiency'] = df_ot_by_day['Earned Hours'] / df_ot_by_day['Hours']


''' PERFORM SOME STATISTICAL ANALYSIS ON THE df AND df_by_day

What does the median saturday look like and how does that compare to the median weekday
    for hours/weight/earned hours

'''




#%%








def plot_hist_of_value_stacked_by_day_of_week(series_name, df, normalized_ot=False):
    normalized = ''
    if normalized_ot:
        normalized = ' (OT=1.5x)'
    
    day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    list_of_list_by_day = []
    
    for day in range(0,7):
        this_day = np.array(df[df['Day of Week'] == day][series_name])
        
        list_of_list_by_day.append(this_day)
        
    plt.figure()
    plt.hist(list_of_list_by_day, stacked=True, bins=10)
    plt.legend(day_of_week)
    plt.title('Histogram of ' + series_name +', by day of week' + normalized)
    plt.show()


# plot_hist_of_value_stacked_by_day_of_week('Hours', df)
# plot_hist_of_value_stacked_by_day_of_week('Weight', df)
# plot_hist_of_value_stacked_by_day_of_week('Earned Hours', df)
# plot_hist_of_value_stacked_by_day_of_week('TTL Efficiency', df)






def plot_hist_of_value_by_specific_weekdayint(series_name, df, weekdays, normalized_ot=False):
    normalized = ''
    if normalized_ot:
        normalized = ' (OT=1.5x)'    
    x = do_day_stuff(series_name, df, weekdays)
    this_day = x['series']
    days = x['days']
    plt.figure()
    plt.hist(this_day, bins=10)
    plt.title('Histogram of ' + series_name + ' on ' + days + normalized)
    plt.show()

# plot_hist_of_value_by_specific_weekdayint('Hours', df, 5)
# plot_hist_of_value_by_specific_weekdayint('Weight', df, 5)
# plot_hist_of_value_by_specific_weekdayint('Earned Hours', df, 5)
# plot_hist_of_value_by_specific_weekdayint('TTL Hours per Ton', df, 5)
# plot_hist_of_value_by_specific_weekdayint('Hours', df, [0,1,2,3,4])
# plot_hist_of_value_by_specific_weekdayint('Weight', df, [0,1,2,3,4])
# plot_hist_of_value_by_specific_weekdayint('Earned Hours', df, [0,1,2,3,4])
# plot_hist_of_value_by_specific_weekdayint('TTL Hours per Ton', df, [0,1,2,3,4])



def plot_boxplot_comparing_all_days_weekdays_saturdays(series_name, df, ylim, remove_outliers=True, normalized_ot=False):
    normalized = ''
    if normalized_ot:
        normalized = ' (OT=1.5x)'
        
    percentile =0.9999
    if remove_outliers:
        percentile = 0.8
    
    everything = remove_outliers_individually(df[series_name], percentile)
    saturdays = remove_outliers_individually(df[df['Day of Week'] == 5][series_name], percentile)
    weekdays = remove_outliers_individually( df[df['Day of Week'] < 5][series_name], percentile)
    
    labels = ['All','Weekdays','Saturdays']
    counts = [str(np.shape(aray)[0]) for aray in [everything, weekdays, saturdays]]
    newlabels = [lbl + '\n(' + i +')' for lbl,i in zip(labels,counts)]
    
    plt.figure()
    plt.violinplot([everything, weekdays, saturdays], showmedians=True)
    plt.ylabel(series_name)
    plt.xticks([1,2,3],newlabels)
    plt.title(series_name + ' for: All Days / Weekdays / Saturdays' + normalized)
    plt.ylim(ylim)
    plt.grid(axis='y')
    plt.show()
    

plot_boxplot_comparing_all_days_weekdays_saturdays('Hours', df, (100,550))
plot_boxplot_comparing_all_days_weekdays_saturdays('Hours', df_ot, (100,550), normalized_ot=True)

plot_boxplot_comparing_all_days_weekdays_saturdays('TTL Efficiency', df, (0,1.2))
plot_boxplot_comparing_all_days_weekdays_saturdays('TTL Efficiency', df_ot, (0,1.2), normalized_ot=True)








def plot_boxplot_of_each_day(series_name, df, ylim, remove_outliers='individual', normalized_ot=False):
    normalized = ''
    if normalized_ot:
        normalized = ' (OT=1.5x)'    
    day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    
    
    keep_this_percentage = 0.90
    z_score_for_percentile = stats.norm.ppf(keep_this_percentage)

    
    if remove_outliers == 'group':
        df = df[np.abs(stats.zscore(df[series_name])) < z_score_for_percentile]
        everything = df[series_name]
        days_arrays = [everything]
        for i,day in enumerate(day_of_week):
            this_day = df[df['Day of Week'] == i][series_name]
            days_arrays.append(this_day)     
            
    elif remove_outliers == 'individual':
        everything = remove_outliers_individually(df[series_name], keep_this_percentage)
        days_arrays = [everything]
        for i,day in enumerate(day_of_week):
            this_day = remove_outliers_individually(df[df['Day of Week'] == i][series_name], keep_this_percentage)
            days_arrays.append(this_day)
    
    else:
        everything = df[series_name]
        days_arrays = [everything]
        for i,day in enumerate(day_of_week):
            this_day = df[df['Day of Week'] == i][series_name]
            days_arrays.append(this_day)

    labels = ['All'] + day_of_week
    labels = [s[:3] for s in labels]
    counts = [str(np.shape(aray)[0]) for aray in days_arrays]
    newlabels = [lbl + '\n(' + i + ')' for lbl,i in zip(labels,counts)]
    
    plt.figure()
    plt.violinplot(days_arrays, showmedians=True)
    plt.ylabel(series_name)
    plt.xticks(list(range(1,len(days_arrays)+1)),newlabels)
    plt.title(series_name + ' by day of week' + normalized)
    plt.grid(axis='y')
    plt.ylim(ylim)
    plt.show()


plot_boxplot_of_each_day('Hours', df, (100,550))
plot_boxplot_of_each_day('Hours', df_ot, (100,550), normalized_ot=True)

# plot_boxplot_of_each_day('Tons', df)
# plot_boxplot_of_each_day('Earned Hours', df)
plot_boxplot_of_each_day('TTL Hours per Ton', df, (0,60))
plot_boxplot_of_each_day('TTL Hours per Ton', df_ot, (0,60), normalized_ot=True)

plot_boxplot_of_each_day('TTL Efficiency', df, (0,1.5))
plot_boxplot_of_each_day('TTL Efficiency', df_ot, (0,1.5), normalized_ot=True)


def plot_reg_ot_norm_next_for_TFS(series_name, df, df_ot, weekday_num, ylim, remove_outliers='individual'):
    day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    day_name = day_of_week[weekday_num]
    
    keep_this_percentage = 0.9
    z_score_for_percentile = stats.norm.ppf(keep_this_percentage)
    
    days_arrays = []
    
    this_day = remove_outliers_individually(df[df['Day of Week'] == weekday_num][series_name], keep_this_percentage)
    this_day_ot = remove_outliers_individually(df_ot[df_ot['Day of Week'] == weekday_num][series_name], keep_this_percentage)
    days_arrays = [this_day, this_day_ot]        
    
    # for i in range(len(days_of_week)):
    #     this_day = remove_outliers_individually(df_ot[df_ot['Day of Week'] == i+4][series_name], keep_this_percentage)
    #     days_arrays.append(this_day)
    
    
    labels = [day_name, day_name + '\n(OT=1.5)']
    counts = [str(np.shape(aray)[0]) for aray in days_arrays]
    newlabels = [lbl + '\n(' + i + ')' for lbl,i in zip(labels,counts)]
    plt.figure()
    plt.violinplot(days_arrays, showmedians=True, widths=0.75)
    plt.ylabel(series_name)
    plt.xticks(list(range(1,len(days_arrays)+1)),newlabels)
    plt.title(series_name + ' (Regular vs. Normalized OT) on ' + day_name + "'s")
    plt.grid(axis='y')
    # plt.ylim(ylim)
    plt.show()    
    

plot_reg_ot_norm_next_for_TFS('Hours', df, df_ot, 4,(100,550), remove_outliers='individual')
plot_reg_ot_norm_next_for_TFS('Hours', df, df_ot, 5,(100,550), remove_outliers='individual')

plot_reg_ot_norm_next_for_TFS('TTL Efficiency', df, df_ot, 4,(0,1.5), remove_outliers='individual')
plot_reg_ot_norm_next_for_TFS('TTL Efficiency', df, df_ot, 5,(0,1.5), remove_outliers='individual')

    
    


    








































#%%
''' Numerical Analysis to compare saturdays to weekdays '''

def return_median_quartiles_for_series(series_name, df, weekdays):
    x = do_day_stuff(series_name, df, weekdays)
    this_day = x['series']
    days = x['days']
        
    
    
    q1, median, q3 = np.quantile(this_day, [0.25,0.5,0.75])
    
    iqr = q3 - q1
    
    uav = q3 + 1.5 * iqr
    lav = q1 - 1.5 * iqr
    
    print('Statistics of ' + series_name + ' on ' + days)
    
    print('UAV:\t{}'.format(np.round(uav,3)))
    print('Q3:\t\t{}'.format(np.round(q3,3)))
    print('Median:\t{}'.format(np.round(median,3)))
    print('Q1:\t\t{}'.format(np.round(q1,3)))
    print('LAV:\t{}'.format(np.round(lav,3)))
    
return_median_quartiles_for_series('TTL Efficiency', df, 5)
return_median_quartiles_for_series('TTL Efficiency', df, [0,1,2,3,4,5])































