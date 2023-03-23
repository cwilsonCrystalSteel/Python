# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 19:37:10 2023

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')
sys.path.append('C:\\Users\\cwilson\\documents\\python\\Speedo_Dashboard')
from Pull_Fablisting_data import get_fablisting_plus_model_summary
from Pull_TimeClock_data import get_timeclock_summary
from Post_to_GoogleSheet import post_observation, post_predictor
import datetime
# this will be the controller for automation?

state = 'TN'
sheet = 'CSM QC Form' 

now = datetime.datetime.now()


now_dt = now.replace(hour=6, minute=0, second=0, microsecond=0)
start_dt = now_dt
# get start_dt as most recent monday
while start_dt.weekday() != 6:
    start_dt -= datetime.timedelta(days=1)

end_dt = start_dt + datetime.timedelta(days=6)
end_dt = end_dt.replace(hour=23, minute=59)
print('running the speedo dashboard for {} to {}'.format(start_dt, end_dt))
# start_dt = now
# if start_dt.hour < 6:
#     start_dt = start_dt - datetime.timedelta(days=1)
#     start_dt = start_dt.replace(hour=6, minute=0, second=0, microsecond=0)
    
# else:
#     start_dt = now.replace(hour=6, minute=0, second=0, microsecond=0)

# end_dt = start_dt + datetime.timedelta(days=1)


fablisting_summary = get_fablisting_plus_model_summary(start_dt, end_dt, sheet=sheet)
timeclock_summary = get_timeclock_summary(start_dt, end_dt, state=state, basis=None)

run_for_date = start_dt.strftime('%m/%d/%Y')
gsheet_dict = {'Date':run_for_date}
gsheet_dict.update(fablisting_summary)
gsheet_dict.update(timeclock_summary)

if gsheet_dict['Direct Hours']:
    post_observation(gsheet_dict)
else:
    print('There was no direct horus so I wont post anything')

predictor = None


# if predictor != None:
#     predictor = get_prediction_formula()
#     post_predictor()


# if now.hour == 6:
#     ''' get the previous days shit for the summary tab '''
#     start_dt = now - datetime.timedelta(days=1)
#     start_dt = start_dt.replace(hour=6, minute=0, second=0, microsecond=0)
#     end_dt = start_dt + datetime.timedelta(days=1)
#     fablisting_summary = get_fablisting_plus_model_summary(start_dt, end_dt, sheet=sheet)
#     timeclock_summary = get_timeclock_summary(start_dt, end_dt, state=state, basis=None)
    
#     run_for_date = start_dt.strftime('%m/%d/%Y')
#     gsheet_dict = {'Date':run_for_date}
#     gsheet_dict.update(fablisting_summary)
#     gsheet_dict.update(timeclock_summary)    
#     post_observation(gsheet_dict, isReal=True, sheet_name='Day Summary')