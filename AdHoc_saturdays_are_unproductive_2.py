# -*- coding: utf-8 -*-
"""
Created on Thu Jun 24 09:05:28 2021

@author: CWilson
"""

import pandas as pd
import datetime
import matplotlib.pyplot as plt

day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
# generated via the AdHoc_Absenteeism.py
absenteeism = pd.read_excel("C://users//cwilson//downloads//absenteeism_by_day.xlsx", index_col=0)

mdi_file = 'Z:\CSM\\MDI Documents - 2021\\C1-Efficiency Results_MDI-CSM.xlsx'

df = pd.read_excel(mdi_file, sheet_name='DataEntry')

df = df[~df['Date'].isna()]

df = df[df['Date'] <= datetime.datetime(2021, 6, 12)]

df = df[df.columns[3:18]]

df['Is Work Day'] = df['Direct'] > 100

df = df[df['Is Work Day']]

df = df[['Date','Earned','Direct','DL EFF','Indirect','OT Hrs','Tons','Is Work Day']]

df['Day of Week'] = df['Date'].dt.weekday

df['Total Hours'] = df['Direct'] + df['Indirect']

df['Realized Hours'] = df['Direct'] + df['Indirect'] - df['OT Hrs'] + (1.5 * df['OT Hrs'])

df['Has OT'] = df['OT Hrs'] > 0



df['Regular TTL Eff'] = df['Earned'] / df['Total Hours']

df['Realized TTL Eff'] = df['Earned'] / df['Realized Hours']

df['Total vs Realized Hours'] = df['Realized Hours'] - df['Total Hours']

df['Total vs Realized Eff'] = df['Regular TTL Eff'] - df['Realized TTL Eff']


only_ot = df[df['Has OT']]

days_with_no_prod = only_ot[only_ot['Tons'] < 5]


working_only = df[df['Is Work Day']]

working_bad_prod = working_only[working_only['DL EFF'] < 0.5]

working_zero_prod = working_only[working_only['DL EFF'] == 0]

working_bad_prod_sat = working_bad_prod[working_bad_prod['Day of Week'] >= 5]

working_zero_prod_sat = working_zero_prod[working_zero_prod['Day of Week'] >= 5]

num_weekend_days = sum(working_only['Day of Week'] >= 5)

percent_of_bad_weekend_days = working_bad_prod_sat.shape[0] / num_weekend_days



print('Percet of weekend days less than 50% Eff:\t{:2.1%}'.format(percent_of_bad_weekend_days))

working_bad_prod_week = working_bad_prod[working_bad_prod['Day of Week'] < 5]

num_week_days = sum(working_only['Day of Week'] < 5)

percent_of_bad_week_days = working_bad_prod_week.shape[0] / num_week_days

print('Percet of weekday days less than 50% Eff:\t{:2.1%}'.format(percent_of_bad_week_days))


working_med_prod = working_only[(working_only['DL EFF'] >= 0.5) & (working_only['DL EFF'] < 0.85)]

working_med_prod_sat = working_med_prod[working_med_prod['Day of Week'] >= 5]

working_good_prod = working_only[working_only['DL EFF'] >= 0.85]

working_good_prod_sat = working_good_prod[working_good_prod['Day of Week'] >= 5]

percent_good_weekend_days = working_good_prod_sat.shape[0] / num_weekend_days

print('Percet of weekend days over 85% Eff:\t{:2.1%}'.format(percent_good_weekend_days))

working_good_prod_week = working_good_prod[working_good_prod['Day of Week'] < 5]

percent_of_good_week_days = working_good_prod_week.shape[0] / num_week_days

print('Percet of weekday days over 85% Eff:\t{:2.1%}'.format(percent_of_good_week_days))


#%% plot 2 histograms
# top histogram is weekday effiecency
# bottom histogram is weekend effieency

hist_start = working_only[working_only['DL EFF'] < 1.25]
hist_weekday = hist_start[hist_start['Day of Week'] < 5]['DL EFF']
hist_weekend = hist_start[hist_start['Day of Week'] >= 5]['DL EFF']

bins = [round(0 + i*0.2, 1) for i in range(0,8)]
xlabels = [str(i*100)+'%' for i in bins]

fig, axes = plt.subplots(nrows=2, ncols=1, sharey=True)
fig.tight_layout()
ax1 = plt.subplot(211)
ax1.hist(hist_weekday, rwidth=0.8, bins=bins)
ax1.set_xticks(bins)
ax1.set_xticklabels(xlabels)
ax1.set_title('Distribution of DL Eff. on Weekdays')
ax2 = plt.subplot(212) 
ax2.hist(hist_weekend, rwidth=0.8, bins=bins)
ax2.set_xticks(bins)
ax2.set_xticklabels(xlabels)
ax2.set_yticks([0,5,10])
ax2.set_title('Distribution of DL Eff. on Weekend Days')

plt.show()



#%% Plot bar chart of abseenteeism by day

bar_start = absenteeism.groupby(by=['Weekday']).sum()
bar_vals = bar_start['Absent'] / (bar_start['Absent'] + bar_start['Present'])
ytickvals = [round(0+0.05*i,2) for i in range(0,10)]
yticklabels = [str(i*100)+'%' for i in ytickvals]


xlabels = [i[:3] for i in day_of_week]
plt.bar(xlabels, bar_vals)
plt.yticks(ytickvals, yticklabels)
plt.grid('major', axis='y')
plt.title('% Absenteeism by Day of Week')



























