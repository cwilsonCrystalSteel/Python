# -*- coding: utf-8 -*-
"""
Created on Wed Jun  1 10:40:12 2022

@author: CWilson
"""
import pandas as pd
import datetime 
import numpy as np
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet


start_date = '01/02/2022'
end_date = '05/21/2022'
csm = grab_google_sheet('CSM QC Form', start_date, end_date)
csf = grab_google_sheet('CSF QC Form', start_date, end_date)
fed = grab_google_sheet('FED QC Form', start_date, end_date)


def fablisting_to_weekly_total(fab_df):
    fab_df['date'] = pd.to_datetime(fab_df['Timestamp']).dt.date
    fab_df_by_day = fab_df.groupby('date').sum().reset_index()
    fab_df_by_day['week'] = np.floor((fab_df_by_day['date'] - datetime.date(2022,1,2)).dt.days /7)
    fab_df_by_week = fab_df_by_day.groupby('week').sum().reset_index()
    fab_df_by_week['tons'] = fab_df_by_week['Weight'] / 2000

    return fab_df_by_week


csm_by_week = fablisting_to_weekly_total(csm)
fed_by_week = fablisting_to_weekly_total(fed)
csf_by_week = fablisting_to_weekly_total(csf)
