# -*- coding: utf-8 -*-
"""
Created on Mon Jul 19 08:15:35 2021

@author: CWilson
"""

from Get_model_estimate_hours_attached_to_fablisting import apply_model_hours
from Grab_Fabrication_Google_Sheet_Data import grab_google_sheet
import pandas as pd


csm = grab_google_sheet('CSM QC Form', '01/01/2020', '08/01/2021')
csm['Shop'] = 'CSM'
csf = grab_google_sheet('CSF QC Form', '01/01/2020', '08/01/2021')
csf['Shop'] = 'CSF'
fed = grab_google_sheet('FED QC Form', '01/01/2020', '08/01/2021')
fed['Shop'] = 'FED'

j2009 = csm[csm['Job #'] == 2009]
j2009 = j2009.append(csf[csf['Job #'] == 2009], ignore_index=True)
j2009 = j2009.append(fed[fed['Job #'] == 2009], ignore_index=True)

j2011 = csm[csm['Job #'] == 2011]
j2011 = j2011.append(csf[csf['Job #'] == 2011], ignore_index=True)
j2011 = j2011.append(fed[fed['Job #'] == 2011], ignore_index=True)



j2009_new = apply_model_hours(j2009)
j2011_new = apply_model_hours(j2011)



# j2009_old = apply_model_hours(j2009)
# j2011_old = apply_model_hours(j2011)



eva_sums = pd.DataFrame(columns=[2009,2011], index=['Model Hours','Hours Worked','Tons'])
eva_sums.loc['Model Hours',2009] = j2009_new.sum()['Earned Hours']
eva_sums.loc['Model Hours',2011] = j2011_new.sum()['Earned Hours']


# CSM + FED + DEL
h2009 = 10171.7 + 4190.45 + 1422.12
h2011 = 21287.9+ 0 + 8

eva_sums.loc['Hours Worked',2009] = h2009
eva_sums.loc['Hours Worked',2011] = h2011

t2009 = j2009.sum()['Weight'] / 2000
t2011 = j2011.sum()['Weight'] / 2000

eva_sums.loc['Tons',2009] = t2009
eva_sums.loc['Tons',2011] = t2011

no_eva2009 = j2009_new[j2009_new['Earned Hours'].isna()]
no_eva2011 = j2011_new[j2011_new['Earned Hours'].isna()]
