# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 13:55:08 2021

@author: CWilson
"""

import pandas as pd
import datetime
from Read_Group_hours_HTML import output_each_clock_entry_job_and_costcode

day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

mdi_file = 'Z:\CSM\\MDI Documents - 2021\\C1-Efficiency Results_MDI-CSM.xlsx'

df = pd.read_excel(mdi_file, sheet_name='DataEntry')

df = df[~df['Date'].isna()]

df = df[df.columns[3:18]]

df['Is Work Day'] = df['Direct'] > 100


group_hours_file = "c:\\users\\cwilson\downloads\\group_hours_YTD_06-12-2021.html"

clocks = output_each_clock_entry_job_and_costcode(group_hours_file)

# clocks = clocks[clocks['Job'] != '648 - Unpaid time off request']
# clocks = clocks[clocks['Job'] != '650 - PTO (Shop)']


employees = list(pd.unique(clocks['Name']))

employees.remove("Jeremy McBride")
employees.remove("Richard Ray")
employees.remove("Zack Harville")
employees.remove('Robert Richardson')
employees.remove('Chester Davis')
employees.remove('Robert Burrell')
employees.remove('Rodger Grant')






working_days = df[['Date','Direct','Is Work Day']]
working_days = working_days.set_index('Date')

x = {}
l = 0
for day in working_days.index[:-10]:
    
    if working_days.loc[day]['Is Work Day']:
        x[day] = []
        l += 1
        for employee in employees:
            chunk = clocks[clocks['Name'] == employee]
            # get only values with dates greater then the day at hand
            chunk = chunk[chunk['Start'] > day]
            # get only values with the dates less than midnight of the day at hand
            chunk = chunk[chunk['Start'] < day + datetime.timedelta(hours=23, minutes=59)]
            
            if chunk.shape[0] == 0:
                x[day].append(employee)

print(l)
i = 0

for date in x.keys():
    i += len(x[date])
print("Number of Days Missed:\t{}".format(i))

# the number of people in the list of employees times the working days
j = len(employees) * working_days[working_days['Is Work Day']].shape[0]
print("Total Number of days that should have been worked:\t{}".format(j))

absenteeism_rate = i/j
print("Absenteeism Rate:\t{:2.2%}".format(absenteeism_rate))



# number of saturdays that are missed

absent = {}
present = {}


for i, dayname in enumerate(day_of_week):
    absent[i] = {}
    present[i] = {}
    
    for day in working_days.index[:-9]:
        if working_days.loc[day]['Is Work Day']:
            if day.weekday() == i:
                absent[i][day] = []
                present[i][day] = []
                # print(day)

                for employee in employees:
                    chunk = clocks[clocks['Name'] == employee]
                    chunk = chunk[chunk['Start'] > day]
                    chunk = chunk[chunk['Start'] < day + datetime.timedelta(hours=23, minutes=59)]
                    if chunk.shape[0] == 0:
                        # print(employee)
                        absent[i][day].append(employee)

                    else:
                        present[i][day].append(employee)

                    

dfs = {}
for i, dayname in enumerate(day_of_week):
    absent_vs_present = {}
    for day in present[i].keys():
        num_absent = len(absent[i][day])
        num_present = len(present[i][day])
        absent_vs_present[day] = {"Absent":num_absent,
                                  "Present":num_present}
    dfs[i] = pd.DataFrame().from_dict(absent_vs_present, orient='index')
    dfs[i]['Precent Absent'] = dfs[i]['Absent'] / (dfs[i]['Absent'] + dfs[i]['Present'])
  
    
big_df = pd.DataFrame()
for i in dfs.keys():
    big_df = big_df.append(dfs[i])
big_df = big_df.sort_index()
big_df['Weekday'] = big_df.index.weekday
big_df.to_excel("C://users//cwilson//downloads//absenteeism_by_day.xlsx")


absenteeism_by_weekday = pd.DataFrame(columns=['# Working Days','Absent','Present','Absenteeism Rate'])
for i, dayname in enumerate(day_of_week):
    summed = dfs[i].sum()
    sum_absent = summed['Absent']
    sum_present = summed['Present']
    average_rate = sum_absent / (sum_absent + sum_present)
    append_dict = {'# Working Days':dfs[i].shape[0], 'Absent':sum_absent, 'Present':sum_present, 'Absenteeism Rate':average_rate}
    absenteeism_by_weekday.loc[dayname] = append_dict
    print(dayname + "'s\t average absenteetism rate is:\t{:2.1%}".format(average_rate))



weekday_absent = 0
weekday_present = 0
for i,dayname in enumerate(day_of_week[:5]):
    summed = dfs[i].sum()
    weekday_absent += summed['Absent']
    weekday_present += summed['Present']

weekday_absenteeism = weekday_absent / (weekday_absent + weekday_present)

saturday_absent = dfs[5].sum()['Absent']
saturday_present = dfs[5].sum()['Present']
saturday_absenteeism = saturday_absent / (saturday_absent + saturday_present)



print('Number of weekdays missed:\t\t\t\t{}'.format(weekday_absent))
print('Number of weekdays worked:\t\t\t\t{}'.format(weekday_present))


print('Number of Saturdays missed:\t\t\t\t{}'.format(saturday_absent))
print('Number of Saturdays worked:\t\t\t\t{}'.format(saturday_present))
print('Weekday  average absenteeism rate is:\t{:2.1%}'.format(weekday_absenteeism))
print('Saturday average absenteeism rate is:\t{:2.1%}'.format(saturday_absenteeism))









