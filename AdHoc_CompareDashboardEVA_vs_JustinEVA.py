# -*- coding: utf-8 -*-
"""
Created on Mon Apr 24 07:38:44 2023

@author: CWilson
"""

import datetime 
import pandas as pd
import numpy as np
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2


''' compare the csv of FORM RESPONSES from fablisting to with_model '''


#%% my way

start_dt = datetime.datetime(2023, 4, 16, 6, 0)
end_dt = datetime.datetime(2023, 4, 22, 23, 59)
sheet= 'CSM QC Form'

start_date = start_dt.strftime('%m/%d/%Y')
# format as the date string
end_date = end_dt.strftime('%m/%d/%Y')
print('Pulling fablisting for: {} to {}'.format(start_date, end_date))
# get fablisting for all of start_date and all of end_date
fablisting = grab_google_sheet(sheet, start_date, end_date)
# get dates between yesterday at 6 am and today at 6 am
fablisting = fablisting[(fablisting['Timestamp'] > start_dt) & (fablisting['Timestamp'] < end_dt)]
# get the model hours attached to fablisting
with_model = apply_model_hours2(fablisting, how = 'model but Justins dumb way of getting average hours', fill_missing_values=True, shop=sheet[:3])
# with_model = apply_model_hours2(fablisting, how = 'model', fill_missing_values=False, shop=sheet[:3])

# get rid of shit I dont' need for this analysis
earned_hours = np.round(with_model['Earned Hours'].sum(), 2)
tonnage = np.round((with_model['Weight'].sum() / 2000), 2)
quantity = int(with_model['Quantity'].sum())


dashboard_output = {'Earned Hours':earned_hours, 'Tons':tonnage, 'Quantity Pieces':quantity}


with_model['Date'] = with_model['Timestamp'].dt.date

model_by_date = with_model.groupby('Date').sum()



#%% from fablisting via downloaded csv

filepath = 'C:\\Users\\cwilson\\Downloads\\Fabrication Listing - QC Input - Form Responses.csv'

justin = pd.read_csv(filepath, header=1)
justin = justin[justin['Site'] == 'CSM']
justin['Date'] = pd.to_datetime(justin['Date'])
justin = justin[(justin['Date'] >= start_dt) & (justin['Date'] <= end_dt)]



#%% join the dfs

with_model['Job #'] = with_model['Job #'].astype(np.int64)
justin['Job #'] = justin['Job #'].astype(np.int64)

joined = pd.merge(with_model, justin, how='left', left_on=['Job #','Lot #','Piece Mark - REV'], right_on=['Job #','Lot #','Piecemark'])
valuable = joined[['Timestamp','Job #','Lot #','Piece Mark - REV','Piecemark','Earned Hours','EVA']]


mine_no_hours = valuable[valuable['Earned Hours'].isna() & valuable['EVA'] > 0]

job_lots = mine_no_hours.groupby(['Job #','Lot #']).agg({'Piecemark':len, 'EVA':sum, 'Piece Mark - REV':lambda x: ', '.join(x)})


