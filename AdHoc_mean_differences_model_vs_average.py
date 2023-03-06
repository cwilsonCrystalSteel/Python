# -*- coding: utf-8 -*-
"""
Created on Mon Aug  2 13:33:09 2021

@author: CWilson
"""

from scipy import stats
import pandas as pd
import numpy as np
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours


start_date = "01/01/2021"
end_date = "07/31/2021"

path = 'c:\\users\\cwilson\\downloads\\'

# test at 95%
alpha = 0.05

# testing_type = '1 Sample on Hours per Ton'
testing_type = '2 Sample on Earned Hours'



shops = ['CSM QC Form', 'FED QC Form','CSF QC Form']
# shops = ['CSM QC Form']



for shop in shops:
    if shop == shops[0]:
        fablisting = grab_google_sheet(shop, start_date, end_date)
        fablisting['Shop'] = shop[:3]
        fablisting = fablisting[['Timestamp','Job #','Lot #','Quantity','Piece Mark - REV', 'Weight','Shop']]
        
    else:
        this_shop = grab_google_sheet(shop, start_date, end_date)
        this_shop['Shop'] = shop[:3]
        this_shop = this_shop[['Timestamp','Job #','Lot #','Quantity','Piece Mark - REV', 'Weight','Shop']]
        fablisting = fablisting.append(this_shop, ignore_index=True)

fablisting = fablisting[~(fablisting['Lot #'] == '')]
fablisting = fablisting[fablisting['Weight'] > 0]

if testing_type == '1 Sample on Hours per Ton':
    table = pd.DataFrame(columns=['Shop','Job #','Lot #','n','mean','stddev','H0','probability','Outcome'])
elif testing_type == '2 Sample on Earned Hours':
    table = pd.DataFrame(columns=['Shop','Job #','Lot #',
                                  'n1','mean1','stddev1',
                                  'n2','mean2','stddev2',
                                  'df','probability','Outcome'])

for shop in shops:
    
    ''' read the averages & get rid of duplicates for other shops '''
    averages = pd.read_excel('c:\\downloads\\averages.xlsx')
    
    duplicates = averages[averages['Job'].duplicated(keep=False)]
    # drop all of the duplicates from the averages df
    averages = averages.drop(index=duplicates.index)
    
    # go thru each job to see if the shop is present for that job
    for job in pd.unique(duplicates['Job']):
        # only get that job from the duplicates df
        chunk = duplicates[duplicates['Job'] == job]
        # only get the ones for that shop
        chunk = chunk[chunk['Shop'] == shop[:3]]
        # if the chunk has zero rows, take the minimum value
        if chunk.shape[0] == 0:
            # grab the duplicates for that job again
            chunk = duplicates[duplicates['Job'] == job]
            # keep the smallest hours per ton
            chunk = chunk.drop_duplicates(subset='Job', keep='first')
        
        # put the chunk back into the averages df
        averages = averages.append(chunk)
            
    averages = averages.set_index('Job')    
    
    
    ''' get the fablisting data '''
    
    this_shop = fablisting[fablisting['Shop'] == shop[:3]]
    
    # get all the different job numbers
    jobs = pd.unique(this_shop['Job #'])
    
    hypothesis_testing = {}
    
    proof_holder = {}
    
    for job in jobs:
    
        this_job = this_shop[this_shop['Job #'] == job]
        
        lots = list(pd.unique(this_job['Lot #']))
        # if '' in lots:
        #     lots.remove('')
        
       
        
        # lot = lots[1]
            
        for lot in lots:
            
            this_lot = this_job[this_job['Lot #'] == lot]
            
            model = apply_model_hours(this_lot, how='model', fill_missing_values=False)
            
            model = model[~model['Earned Hours'].isna()]
            
            
            # really needs to be sample size of >30 for both
            # only the model will have less samples then the old way            
            if model.shape[0] < 25:
                print('Sample size too small ' + str(job) + ' lot ' + str(lot))
                # table = table.append({'Shop': shop,
                #                       'Job #': job,
                #                       'Lot #': lot,
                #                       'n': model.shape[0]},
                #                      ignore_index=True)
                continue
            
            
            if testing_type == '2 Sample on Earned Hours':
                
                model_series = model['Earned Hours']
                
                old = apply_model_hours(this_lot, how='old way', shop=shop[:3])
                
                old_series = old['Earned Hours']
                
                # let the model be #1
                mean1 = model_series.mean()
                std1 = stats.tstd(model_series)
                n1 = model_series.shape[0]
                
                # let the old way be #2
                mean2 = old_series.mean()
                std2 = stats.tstd(old_series)
                n2 = old_series.shape[0]    
                
                print('Job: \t' + str(job))
                print('Lot: \t' + str(lot))
                print('Sample Sizes:')
                print('Model: \t' + str(n1))
                print('Old: \t' + str(n2))
                
                
                # F-test to determine variance equality
                f_test = std1**2 / std2**2
                nu1 = n1 - 1
                nu2 = n2 - 1
                # probablity of P(F<F_test)
                p_f_less = stats.f.cdf(f_test, dfn=nu1, dfd=nu2)
                p_f = 2 * np.min([p_f_less, 1-p_f_less])
                
                
                if p_f < alpha:
                    print('Reject variance equality hypothesis')
                    # unknown & not equal variance t-test
                    t_test = (mean1 - mean2) / (std1**2/n1 + std2**2/n2)**0.5
                    nu = (std1**2/n1 + std2**2/n2)**2 / ((std1**2/n1)**2/(n1-1) + (std2**2/n2)**2/(n2-1))
                    nu = np.floor(nu)
        
                elif p_f > alpha:
                    print('Fail to reject variance equality hypothesis')
                    # unknown & equal variances t-test
                    std_pooled = (nu1 * std1**2 + nu2 * std2**2) / (nu1 + nu2)
                    t_test = (mean1 - mean2) / (std_pooled * (1/n1 + 1/n2)**0.5)
                    nu = nu1 + nu2
            
                
                # probablity of P(t<t_test)
                p_t_less = stats.t.cdf(t_test, nu)
                p_t = 2 * np.min([p_t_less, 1 - p_t_less])             

                table = table.append({'Shop': shop[:3],
                                      'Job #': job,
                                      'Lot #': lot,
                                      'n1': n1,
                                      'mean1': mean1,
                                      'stddev1': std1,
                                      'n2': n2,
                                      'mean2': mean2,
                                      'stddev2': std2,
                                      'df': nu,
                                      'probability': p_t},
                                     ignore_index=True)
            
            
            elif testing_type == '1 Sample on Hours per Ton':
                
                model_series = model['Earned Hours'] / (model['Weight'] / 2000)
                
                mean0 = averages.loc[job, 'Hours per Ton']
                
                mean1 = model_series.mean()
                std1 = stats.tstd(model_series)
                n1 = model_series.shape[0]
                nu = n1 - 1
                
                t_test = (mean1 - mean0) / (std1 / n1**0.5)
            
            
            
            
                # probablity of P(t<t_test)
                p_t_less = stats.t.cdf(t_test, nu)
                p_t = 2 * np.min([p_t_less, 1 - p_t_less])            
            
            
            
                table = table.append({'Shop': shop[:3],
                                      'Job #': job,
                                      'Lot #': lot,
                                      'n': n1,
                                      'mean': mean1,
                                      'stddev': std1,
                                      'H0': mean0,
                                      'probability': p_t},
                                     ignore_index=True)
            
            
            
            
            # if the job has not been created yet, create it in the dict
            if hypothesis_testing.get(job) == None:
                hypothesis_testing[job] = {'Tests':1,
                                           'Reject Equality': 0,
                                           'Unequal Means': [],
                                           'Fail to Reject Equality': 0,
                                           'Equal Means': []}   

            # just add to the test counter if it has been created
            else:
                hypothesis_testing[job]['Tests'] += 1
            
            
            if p_t > alpha:
                print('Fail to reject hypothesis of mean equality')
                hypothesis_testing[job]['Fail to Reject Equality'] += 1
                hypothesis_testing[job]['Equal Means'].append(lot)
                
            elif p_t < alpha:
                print('Reject hypothesis of mean equality')
                hypothesis_testing[job]['Reject Equality'] += 1
                hypothesis_testing[job]['Unequal Means'].append(lot)
                
                if proof_holder.get(job) == None:
                    proof_holder[job] = {}
                
                if testing_type == '1 Sample on Hours per Ton':
                    proof_holder[job][lot] = model.sort_values(by='Earned Hours')
                    
                elif testing_type == '2 Sample on Earned Hours':
                    proof_holder[job][lot] = old.join(model['Earned Hours'], rsuffix=' (New)').sort_values(by='Earned Hours')
            
    
    
    
    
    summary = pd.DataFrame().from_dict(hypothesis_testing)
    
  
    
    
    # with pd.ExcelWriter(path + shop[:3] + ' ' + testing_type + ' mean difference testing' + '.xlsx') as writer:
    #         summary.to_excel(writer, sheet_name='Summary')
            
    #         for job in proof_holder.keys():
    #             for lot in proof_holder[job].keys():
    #                 sheet_name = str(job) + ' Lot ' + lot
    #                 proof_holder[job][lot].to_excel(writer, sheet_name=sheet_name)



# set the outcome column to be False if probability < alpha
table['Outcome'] = ~ (table['probability'] < alpha )

if 'mean1' in table.columns:
    table = table.sort_values(by=['Job #','Outcome','mean1'])
else:
    table = table.sort_values(by=['Job #','Outcome','mean'])

table = table.reset_index(drop=True)

with pd.ExcelWriter(path + 'Summary data ' + testing_type + ' mean difference testing' + '.xlsx') as writer:
    table.to_excel(writer, sheet_name='Summarized Values')
