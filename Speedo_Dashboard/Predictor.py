# -*- coding: utf-8 -*-
"""
Created on Sat Mar 11 08:53:58 2023

@author: CWilson
"""
import sys
sys.path.append("C:\\Users\\cwilson\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python39\\site-packages")
sys.path.append('C:\\Users\\cwilson\\documents\\python')

from sklearn.linear_model import LinearRegression, Ridge, Lasso
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.metrics import mean_squared_error, make_scorer
from sklearn.preprocessing import MinMaxScaler
import datetime
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
# this will be the function that does prediction 
now = datetime.datetime.now()


def get_prediction_dict(df):
    
    
    return None



def smape(y_test, y_pred):
    numerator = np.abs(y_test-y_pred)
    denominator = (y_test + np.abs(y_pred)) /200
    return np.mean(np.divide(numerator,denominator))

smape_scorer = make_scorer(smape, greater_is_better=False)


for csv in ['CSM', 'CSF', 'FED']:
    filepath = 'C:\\Users\\cwilson\\Documents\\Python\\Speedo_Dashboard\\Archive_' + csv + '.csv'
    temp_df = pd.read_csv(filepath)
    temp_df['Shop'] = csv
    if csv == 'CSM':
        df = temp_df
    else:
        df = pd.concat([df, temp_df.reset_index(drop=True)])
# training_csv = 'C:\\Users\\cwilson\\Documents\\Python\\Speedo_Dashboard\\Archive_CSM.csv'
# df = pd.read_csv(training_csv)

df = df.rename(columns={'Date': 'StartOfWeek'})
# df = df[df['StartOfWeek'] == '08/20/2023']
df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
df = df[~df['Timestamp'].isna()]
df['StartOfWeek'] = pd.to_datetime(df['StartOfWeek'])
df['EndOfWeek'] = df['StartOfWeek'] + datetime.timedelta(days=6,hours=23,minutes=59)
df['PercentageOfWeek'] = ((7 * 24 * 60 * 60) - (df['EndOfWeek'] - df['Timestamp']).dt.total_seconds()) / (7 * 24 * 60 * 60)
df = df[~((df['Earned Hours'] == 0) & (df['Direct Hours'] == 0))]
df = df.sort_values(by=['Shop','Timestamp'])
X = df[['PercentageOfWeek', 'Direct Hours','Indirect Hours', 'Number Employees', 'Tons','Quantity Pieces']].to_numpy()
y = df['Earned Hours'].to_numpy()


# %%

shops = pd.unique(df['Shop'])
for i in shops:
    subset = df[df['Shop'] == i]
    plt.plot(subset['Timestamp'].to_numpy(), subset['Earned Hours'].to_numpy())

plt.legend(shops)




# %%



X_train, X_test, y_train, y_test = train_test_split(X, y)

reg = LinearRegression()
reg.fit(X_train, y_train)
reg_pred = reg.predict(X_test)
reg_mse = mean_squared_error(y_test, reg_pred)
reg_smape = smape(y_test, reg_pred)

mse_scorer = make_scorer(mean_squared_error, greater_is_better=False)

ridge = GridSearchCV(Ridge(), {'alpha':np.arange(500,1000, 1)}, scoring=smape_scorer, cv=5)
ridge.fit(X_train, y_train)
ridge_pred = ridge.predict(X_test)
ridge_alpha = ridge.best_params_['alpha']
ridge_mse = mean_squared_error(y_test, ridge_pred)
ridge_smape = smape(y_test, ridge_pred)

lasso = GridSearchCV(Lasso(), {'alpha':np.arange(0,1,0.01)}, scoring=mse_scorer, cv=5)
lasso.fit(X_train, y_train)
lasso_pred = lasso.predict(X_test)
lasso_alpha = lasso.best_params_['alpha']
lasso_mse = mean_squared_error(y_test, lasso_pred)
lasso_smape = smape(y_test, lasso_pred)



gbr=GradientBoostingRegressor()
gbr.fit(X_train, y_train)
gbr_pred = gbr.predict(X_test)
gbr_mse = mean_squared_error(y_test, gbr_pred)
gbr_smape = smape(y_test, gbr_pred)

















