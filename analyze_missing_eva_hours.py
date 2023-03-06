# -*- coding: utf-8 -*-
"""
Created on Thu Jul 14 09:08:20 2022

@author: CWilson
"""


from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours2, fill_missing_model_earned_hours
import pandas as pd
import numpy as np

start_date = '01/01/2021'
end_date = '08/01/2022'

fab_fed = grab_google_sheet('FED QC Form', start_date, end_date)
fab_csf = grab_google_sheet('CSF QC Form', start_date, end_date)
fab_csm = grab_google_sheet('CSM QC Form', start_date, end_date)

keep_cols = ['Timestamp','Job #','Lot #','Quantity','Piece Mark - REV','Weight']
fab_fed = fab_fed[keep_cols]
fab_csf = fab_csf[keep_cols]
fab_csm = fab_csm[keep_cols]


fab_fed_model = apply_model_hours2(fab_fed, fill_missing_values=False, shop='FED')
fab_csf_model = apply_model_hours2(fab_csf, fill_missing_values=False, shop='CSF')
fab_csm_model = apply_model_hours2(fab_csm, fill_missing_values=False, shop='CSM')


fab_fed_model['shop'] = 'FED'
fab_csf_model['shop'] = 'CSF'
fab_csm_model['shop'] = 'CSM'

big_df = pd.concat([fab_fed_model, fab_csf_model, fab_csm_model])

# estimate missing hours in a number of ways
# average eva by piece
def ave_per_piece(df):
    model_only = df[df['Has Model']]
    earned_hours = sum(model_only['Earned Hours'])
    quantity = sum(model_only['Quantity'])
    ave_earned_hours_per_piece = earned_hours / quantity
    return ave_earned_hours_per_piece


ave_per_piece_fed = ave_per_piece(fab_fed_model)
ave_per_piece_csf = ave_per_piece(fab_csf_model)
ave_per_piece_csm = ave_per_piece(fab_csm_model)
ave_per_piece_all = ave_per_piece(big_df)


# average eva per piece & weight per job
def ave_per_job(df):
    model_only = df[df['Has Model']]
    model_only_by_job = model_only.groupby('Job #').sum()[['Quantity','Weight','Earned Hours']]
    model_only_by_job['Ave Earned Hours by Job by Piece'] = model_only_by_job['Earned Hours'] / model_only_by_job['Quantity']
    model_only_by_job['Ave Earned Hours by Job by Weight'] = model_only_by_job['Earned Hours'] / model_only_by_job['Weight']    
    model_only_by_job = model_only_by_job[['Ave Earned Hours by Job by Piece','Ave Earned Hours by Job by Weight']]
    
    piece_col = model_only_by_job['Ave Earned Hours by Job by Piece']
    piece_col = piece_col[~piece_col.isna()]
    piece_col = piece_col[piece_col < np.inf]
    piece_col = piece_col[piece_col > -np.inf]
    ave_piece = sum(piece_col) / len(piece_col)
    
    new_piece_col = model_only_by_job['Ave Earned Hours by Job by Piece'].replace([np.NaN, np.inf, -np.inf], ave_piece)

    
    weight_col = model_only_by_job['Ave Earned Hours by Job by Weight']
    weight_col = weight_col[~weight_col.isna()]
    weight_col = weight_col[weight_col < np.inf]
    weight_col = weight_col[weight_col > -np.inf]
    ave_weight = sum(weight_col) / len(weight_col)    
    
    new_weight_col = model_only_by_job['Ave Earned Hours by Job by Weight'].replace([np.NaN, np.inf, -np.inf], ave_weight)
    
    
    model_only_by_job['Ave Earned Hours by Job by Piece'] = new_piece_col
    model_only_by_job['Ave Earned Hours by Job by Weight'] = new_weight_col    
    
    
    return model_only_by_job


ave_per_job_fed = ave_per_job(fab_fed_model)
ave_per_job_csf = ave_per_job(fab_csf_model)
ave_per_job_csm = ave_per_job(fab_csm_model)
ave_per_job_all = ave_per_job(big_df)


# average eva by weight
def ave_per_weight(df):
    model_only = df[df['Has Model']]
    earned_hours = sum(model_only['Earned Hours'])
    weight = sum(model_only['Weight'])
    ave_earned_hours_per_pound = earned_hours / weight
    return  ave_earned_hours_per_pound

ave_per_weight_fed = ave_per_weight(fab_fed_model)
ave_per_weight_csf = ave_per_weight(fab_csf_model)
ave_per_weight_csm = ave_per_weight(fab_csm_model)
ave_per_weight_all = ave_per_weight(big_df)





missing_all = big_df[~big_df['Has Model']]
missing_all = missing_all[['Timestamp','Job #','Lot #','Quantity','Piece Mark - REV','Weight','Lot Name','shop']]
missing_all['Month'] = missing_all['Timestamp'].dt.month
missing_all['Year'] = missing_all['Timestamp'].dt.year
missing_fed = missing_all[missing_all['shop'] == 'FED']
missing_csf = missing_all[missing_all['shop'] == 'CSF']
missing_csm = missing_all[missing_all['shop'] == 'CSM']


missing_all['Est EVA by Ave Hours per Piece'] = missing_all['Quantity'] * ave_per_piece_all
missing_fed['Est EVA by Ave Hours per Piece'] = missing_fed['Quantity'] * ave_per_piece_fed
missing_csf['Est EVA by Ave Hours per Piece'] = missing_csf['Quantity'] * ave_per_piece_csf
missing_csm['Est EVA by Ave Hours per Piece'] = missing_csm['Quantity'] * ave_per_piece_csm


missing_all['Est EVA by Ave Hours per Weight'] = missing_all['Weight'] * ave_per_weight_all
missing_fed['Est EVA by Ave Hours per Weight'] = missing_fed['Weight'] * ave_per_weight_fed
missing_csf['Est EVA by Ave Hours per Weight'] = missing_csf['Weight'] * ave_per_weight_csf
missing_csm['Est EVA by Ave Hours per Weight'] = missing_csm['Weight'] * ave_per_weight_csm







missing_all = missing_all.merge(ave_per_job_all, how='left', on='Job #')
missing_all['Est EVA by Ave Hours per Piece by Job'] = missing_all['Quantity'] * missing_all['Ave Earned Hours by Job by Piece']
missing_all['Est EVA by Ave Hours per Weight by Job'] = missing_all['Weight'] * missing_all['Ave Earned Hours by Job by Weight']


missing_fed = missing_fed.merge(ave_per_job_fed, how='left', on='Job #')
missing_fed['Est EVA by Ave Hours per Piece by Job'] = missing_fed['Quantity'] * missing_fed['Ave Earned Hours by Job by Piece']
missing_fed['Est EVA by Ave Hours per Weight by Job'] = missing_fed['Weight'] * missing_fed['Ave Earned Hours by Job by Weight']


missing_csf = missing_csf.merge(ave_per_job_csf, how='left', on='Job #')
missing_csf['Est EVA by Ave Hours per Piece by Job'] = missing_csf['Quantity'] * missing_csf['Ave Earned Hours by Job by Piece']
missing_csf['Est EVA by Ave Hours per Weight by Job'] = missing_csf['Weight'] * missing_csf['Ave Earned Hours by Job by Weight']


missing_csm = missing_csm.merge(ave_per_job_csm, how='left', on='Job #')
missing_csm['Est EVA by Ave Hours per Piece by Job'] = missing_csm['Quantity'] * missing_csm['Ave Earned Hours by Job by Piece']
missing_csm['Est EVA by Ave Hours per Weight by Job'] = missing_csm['Weight'] * missing_csm['Ave Earned Hours by Job by Weight']


missing_df_cols_to_drop = ['Ave Earned Hours by Job by Piece','Ave Earned Hours by Job by Weight']
missing_all = missing_all.drop(columns=missing_df_cols_to_drop)
missing_fed = missing_fed.drop(columns=missing_df_cols_to_drop)
missing_csf = missing_csf.drop(columns=missing_df_cols_to_drop)
missing_csm = missing_csm.drop(columns=missing_df_cols_to_drop)





missing_all['Est EVA by Ave Hours per Piece by Job'] = missing_all['Est EVA by Ave Hours per Piece by Job'].fillna(missing_all['Est EVA by Ave Hours per Piece'])
missing_all['Est EVA by Ave Hours per Weight by Job'] = missing_all['Est EVA by Ave Hours per Weight by Job'].fillna(missing_all['Est EVA by Ave Hours per Weight'])

missing_fed['Est EVA by Ave Hours per Piece by Job'] = missing_fed['Est EVA by Ave Hours per Piece by Job'].fillna(missing_fed['Est EVA by Ave Hours per Piece'])
missing_fed['Est EVA by Ave Hours per Weight by Job'] = missing_fed['Est EVA by Ave Hours per Weight by Job'].fillna(missing_fed['Est EVA by Ave Hours per Weight'])

missing_csf['Est EVA by Ave Hours per Piece by Job'] = missing_csf['Est EVA by Ave Hours per Piece by Job'].fillna(missing_csf['Est EVA by Ave Hours per Piece'])
missing_csf['Est EVA by Ave Hours per Weight by Job'] = missing_csf['Est EVA by Ave Hours per Weight by Job'].fillna(missing_csf['Est EVA by Ave Hours per Weight'])

missing_csm['Est EVA by Ave Hours per Piece by Job'] = missing_csm['Est EVA by Ave Hours per Piece by Job'].fillna(missing_csm['Est EVA by Ave Hours per Piece'])
missing_csm['Est EVA by Ave Hours per Weight by Job'] = missing_csm['Est EVA by Ave Hours per Weight by Job'].fillna(missing_csm['Est EVA by Ave Hours per Weight'])







def groupby_month_year(df):
    if 'Month' in df.columns and 'Year' in df.columns:
        by_month = df.groupby(by=['Month','Year']).sum()
        by_month = by_month.drop(columns=['Job #'])
    else:
        df['Month'] = df['Timestamp'].dt.month
        df['Year'] = df['Timestamp'].dt.year
        by_month = groupby_month_year(df)
    
    return by_month


est_missing_all = groupby_month_year(missing_all)
est_missing_fed = groupby_month_year(missing_fed)
est_missing_csf = groupby_month_year(missing_csf)
est_missing_csm = groupby_month_year(missing_csm)


eva_all = groupby_month_year(big_df)[['Quantity','Weight','Earned Hours']]
eva_fed = groupby_month_year(fab_fed_model)[['Quantity','Weight','Earned Hours']]
eva_csf = groupby_month_year(fab_csf_model)[['Quantity','Weight','Earned Hours']]
eva_csm = groupby_month_year(fab_csm_model)[['Quantity','Weight','Earned Hours']]



missing_and_earned_all = eva_all.merge(est_missing_all, on=['Month','Year'], suffixes=(' of Pcs with model',' of Pcs missing model'))
missing_and_earned_fed = eva_fed.merge(est_missing_fed, on=['Month','Year'], suffixes=(' of Pcs with model',' of Pcs missing model'))
missing_and_earned_csf = eva_csf.merge(est_missing_csf, on=['Month','Year'], suffixes=(' of Pcs with model',' of Pcs missing model'))
missing_and_earned_csm = eva_csm.merge(est_missing_csm, on=['Month','Year'], suffixes=(' of Pcs with model',' of Pcs missing model'))


path = 'c:\\users\\cwilson\\documents\\analyze_missing_hours\\analyze missing hours '
path += start_date.replace('/','-') + ' to ' + end_date.replace('/','-')
path += '.xlsx'
with pd.ExcelWriter(path) as writer:
    missing_and_earned_all.to_excel(writer, 'ALL shops summary')
    missing_and_earned_fed.to_excel(writer, 'FED summary')
    missing_and_earned_csf.to_excel(writer, 'CSF summary')
    missing_and_earned_csm.to_excel(writer, 'CSM summary')
    missing_all.to_excel(writer, 'ALL details of missing')
    missing_fed.to_excel(writer, 'FED details of missing')
    missing_csf.to_excel(writer, 'CSF details of missing')
    missing_csm.to_excel(writer, 'CSM details of missing')
    
    
