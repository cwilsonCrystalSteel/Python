# -*- coding: utf-8 -*-
"""
Created on Thu May 13 08:51:50 2021

@author: CWilson
"""

import sys
sys.path.append('c://users//cwilson//documents//python//Attendance Project//')
import pandas as pd
import datetime
import glob
import os
from attendance_google_sheets_credentials_startup import init_google_sheet
import matplotlib.pyplot as plt

# get the production worksheet 
sh = init_google_sheet("1HIpS0gbQo8q1Pwo9oQgRFUqCNdud_RXS3w815jX6zc4")
# get the values from the shipping schedule as a list of lists
worksheet = sh.worksheet('Shipping Sched.').get_all_values()
# convert to a dataframe. row 2 is columns, dont care about stuff before row 10
df = pd.DataFrame(worksheet[10:], columns=worksheet[2])

# empty dict to get rid of line breaks in the column names
new_cols = {}
for col in df.columns:
    new_col = col.replace('\n', ' ')
    new_cols[col] = new_col

# replace columns with new columns w/o line breaks
df = df.rename(columns=new_cols)
# rename the 2nd column which is named '9' for some stupid reason
df = df.rename(columns = {'9':'Seq'})
# convert the shipped column to lower case & keep rows if it does not contain the word yes
df = df[~df['Shipped'].str.lower().str.contains('yes')]
# only get the sequences for CSM
df = df[df['Fabrication Site'] == 'CSM']
# create a copy of the dataframe and rename it 
seq_shipping = df.copy()
# only keep the job, sequence, and delivery date columns
seq_shipping = seq_shipping[['Job','Seq', 'Delivery']]
# get the unique jobs from the dataframe 
seq_shipping_jobs = pd.unique(seq_shipping['Job'])
# convert the delivery date column to datetime
seq_shipping['Delivery'] = pd.to_datetime(seq_shipping['Delivery'], errors='coerce')
# get todays date
today = datetime.datetime.today()
# go XX number of days into the future
future_cutpff_date = today + datetime.timedelta(days=30)
# only keep the stuff due before the future_cutoff_date
seq_shipping = seq_shipping[seq_shipping['Delivery'] <= future_cutpff_date]
# sort by the delivery date, most recent is at the top
seq_shipping = seq_shipping.sort_values(by='Delivery')
# convert the delivery column to just be a date, not a datetime - for excel viewing purposes
seq_shipping['Delivery'] = seq_shipping['Delivery'].dt.date
# get rid of all entries where the sequence contains the word ticket
seq_shipping = seq_shipping[~seq_shipping['Seq'].str.contains('Ticket')]
# the extract('(\d+)') is a regex to get the first instance of numbers out of the string
seq_shipping['NUM_SEQ'] = seq_shipping['Seq'].str.extract('(\d+)')
# combine the job & seq number into one column with '<>' to seperate them
seq_shipping['JOB_SEQ'] = seq_shipping['Job'] + '<>' + seq_shipping['NUM_SEQ']



# get the Lots Log sheet
worksheet = sh.worksheet('LOTS Log').get_all_values()
# covnert it to a dataframe
df = pd.DataFrame(worksheet[1:], columns=worksheet[0])
# empty dict to get rid of line breaks in the column names
new_cols = {}
for col in df.columns:
    new_col = col.replace('\n', ' ')
    new_cols[col] = new_col

# rename the columns with the new columns without line breaks
df = df.rename(columns=new_cols)
# the first column is an empty string but is actually the job column
df = df.rename(columns={'':'Job'})
# only get the CSM lots
df = df[df['Fabrication  Site'] == 'CSM']


# initialize a dataframe 
y = pd.DataFrame(columns=['Job', 'Seq', 'NUM_SEQ', 'LOT Name','Transmittal'])
# keep the Lot Name
y['LOT Name'] = df['LOTS Name']
# Keep the transmittal #
y['Transmittal'] = df['Transmittal #']
# keep the Sequence
y['Seq'] = df['Seq. #']
# keep the Job
y['Job'] = df['Job']
# get just the first instance of numbers from the seq 
y['NUM_SEQ'] = df['Seq. #'].str.extract('(\d+)')
# combine the job & sequence number with the '<>' seperator
y['JOB_SEQ'] = y['Job'] + '<>' + y['NUM_SEQ']
# get the lots not in seq_shipping by JOB_SEQ
lots_not_in_seq_shipping = y[~y['JOB_SEQ'].isin(seq_shipping['JOB_SEQ'])]
# get only the lots that are in the shipping schedule by matching the 'JOB_SEQ' values
y = y[y['JOB_SEQ'].isin(seq_shipping['JOB_SEQ'])]
# set the index of lots_df to be the job-seq
y = y.set_index('JOB_SEQ')
# set the index of the sequence shipping df to be the JOB_SEQ
seq_shipping = seq_shipping.set_index('JOB_SEQ')
# join the delivery date column based on the 'JOB_SEQ' being equal'
y = y.join(seq_shipping['Delivery'])
# sort by the closest delivery dates at top
y = y.sort_values(by=('Delivery'))
# reset the index
y = y.reset_index(drop=True)
# reset the index
seq_shipping = seq_shipping.reset_index(drop=True)
# only keep the columns we want, get rid of 'NUM_SEQ'
y = y[['Job', 'Seq', 'LOT Name', 'Transmittal', 'Delivery']]


# get all instances of duplicated lots except last of the duplicated LOTS
dupe_lots = y[y['LOT Name'].duplicated()]
# get the rows that are not duplicates of lot name
unique_lots = y[~y['LOT Name'].duplicated()]

# This is to get the name of all the sequences present on a LOT
new_seqs = []
# iterate thru each lot that is duplicated 
for lot in dupe_lots['LOT Name']:
    # get those lots from the 
    chunk = y[y['LOT Name'] == lot]
    # get unique of all of the sequences that occur for that lot
    seqs = pd.unique(chunk['Seq']).tolist()
    # join all of the sequences as str for that lot with a comma & space between them
    seqs = ', '.join(seqs)
    # append them to the new_sequences lits
    new_seqs.append(seqs)

# create a copy so that you dont get a setting of copy warning ???????
dupe_lots = dupe_lots.copy()
# set the column to be equal to the 
dupe_lots['Seq'] = new_seqs

# set the index to be the LOT name
unique_lots = unique_lots.set_index('LOT Name')
# set the index to be the LOT name
dupe_lots = dupe_lots.set_index('LOT Name')
# set the unique lots 'Seq' column to be the new seq column in the dupe_lots df based on the index (Lot name)
unique_lots.loc[dupe_lots.index,'Seq'] = dupe_lots['Seq']
# reset the index of the unique lots
unique_lots = unique_lots.reset_index()
# sort the unique lots first by delivery date and then by job 
unique_lots = unique_lots.sort_values(by=['Delivery','Job'])
# get rid of 1926 b/c it already shipped
unique_lots = unique_lots[unique_lots['Job'] != '1926']
# group the unique_lots by the count of delivery date
count_by_date = unique_lots.groupby('Delivery').count()
# only keep the Lot Name column
count_by_date = count_by_date['LOT Name']
# Renmae the series 
count_by_date = count_by_date.rename('# of LOTS Due')
# for excel purposes, only keep the following columns
unique_lots_xlsx = unique_lots[['LOT Name','Job','Seq','Delivery']]
# reset the index so that excel shows a count even tho it starts at 0 and people wont understand that
unique_lots_xlsx = unique_lots_xlsx.reset_index(drop=True)

# the excel file name to write to
xl_file_name = 'c://users/cwilson/downloads/lots_delivery_date.xlsx'
# write to excel with 2 sheets in the file
with pd.ExcelWriter(xl_file_name) as writer:
    unique_lots_xlsx.to_excel(writer, sheet_name='Lots Delivery Date')
    count_by_date.to_excel(writer, sheet_name='# Deliveries by Date')
# all i need to do is go into the excel file and format the 'Lots Delivery Date' to be readabl
# and create a bar chart for # of delvieries by date



