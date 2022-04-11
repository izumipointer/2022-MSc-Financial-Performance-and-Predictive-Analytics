# -*- coding: utf-8 -*-
"""
Created on Wed Jan 12 21:13:34 2022

@author: izumi
"""

#modules data manipulation
import pandas as pd
import numpy as np
import datetime as dt

#modules data visualization
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.seasonal import seasonal_decompose

#data loading csv
Australia=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/Australia.csv', parse_dates=['Date'], index_col='Date')
Canada=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/Canada.csv', parse_dates=['Date'], index_col='Date')
Germany=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/Germany.csv', parse_dates=['Date'], index_col='Date')
Japan=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/Japan.csv', parse_dates=['Date'], index_col='Date')
Mexico=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/Mexico.csv', parse_dates=['Date'], index_col='Date')
Nigeria=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/Nigeria.csv', parse_dates=['Date'], index_col='Date')
UnitedStates=pd.read_csv('C:/Users/izumi/OneDrive/Desktop/Temporary/USA.csv', parse_dates=['Date'], index_col='Date')
UnitedStates['Country']=('USA')

#data loading excel
geo=pd.read_excel('C:/Users/izumi/OneDrive/Desktop/Temporary/MDM_GEO.xlsx')
product=pd.read_excel('C:/Users/izumi/OneDrive/Desktop/Temporary/MDM_PRODUCT.xlsx')
manufacturer=pd.read_excel('C:/Users/izumi/OneDrive/Desktop/Temporary/MDM_MANUFACTURER.xlsx')

#combine worldwide sales
World=pd.concat([Australia, Canada, Germany, Japan, Mexico, Nigeria, UnitedStates])

#==============================================================================

#WW ANALYSIS
#1) Elaborate a tab ranking country by revenue
World.dtypes
World_ranking=World.drop(['ProductID','Units','Zip'], axis=1)
World_ranking=World_ranking.groupby('Country').sum()
World_ranking=World_ranking.sort_values(by='Revenue', ascending=False)
print(World_ranking)

#2) Which countries have the most important revenue growth (in %) between 2011 & 2012?
#prep 2011 df
World_2011=World.drop(['ProductID','Units','Zip'], axis=1)
World_2011.rename(columns={'Revenue': 'Revenue 2011'}, inplace=True)
World_2011=World_2011[(World_2011.index.year==2011)].groupby('Country').sum()

#prep 2012 df
World_2012=World.drop(['ProductID','Units','Zip'], axis=1)
World_2012.rename(columns={'Revenue': 'Revenue 2012'}, inplace=True)
World_2012=World_2012[(World_2012.index.year==2012)].groupby('Country').sum()

#merge 2011 and 2012 dfs and calculate growth %
World_growth=pd.merge(World_2011, World_2012, on='Country')
World_growth['Growth %'] = (World_growth['Revenue 2012'] / World_growth['Revenue 2011'] - 1)
World_growth=World_growth.sort_values(by='Growth %', ascending=False).head(3) #top 3 
print(World_growth)

#3) Which countries have the most important cumulative revenue growth (in %) between 2011 to 2018?

#calc cumulative growth for each country's df
Australia['Year']=Australia.index.year
Australia_cg=Australia.drop(['ProductID', 'Zip', 'Units'], axis=1)
Australia_cg=Australia_cg.groupby('Year').sum()
Australia_cg.rename(columns={'Revenue': 'Australia Revenue'}, inplace=True)
Australia_cg=Australia_cg.pct_change()
Australia_cg = (1 + Australia_cg['Australia Revenue']).cumprod() - 1

Canada['Year']=Canada.index.year
Canada_cg=Canada.drop(['ProductID', 'Zip', 'Units'], axis=1)
Canada_cg=Canada_cg.groupby('Year').sum()
Canada_cg.rename(columns={'Revenue': 'Canada Revenue'}, inplace=True)
Canada_cg=Canada_cg.pct_change()
Canada_cg = (1 + Canada_cg['Canada Revenue']).cumprod() - 1

Germany['Year']=Germany.index.year
Germany_cg=Germany.drop(['ProductID', 'Zip', 'Units'], axis=1)
Germany_cg=Germany_cg.groupby('Year').sum()
Germany_cg.rename(columns={'Revenue': 'Germany Revenue'}, inplace=True)
Germany_cg=Germany_cg.pct_change()
Germany_cg = (1 + Germany_cg['Germany Revenue']).cumprod() - 1

Japan['Year']=Japan.index.year
Japan_cg=Japan.drop(['ProductID', 'Zip', 'Units'], axis=1)
Japan_cg=Japan_cg.groupby('Year').sum()
Japan_cg.rename(columns={'Revenue': 'Japan Revenue'}, inplace=True)
Japan_cg=Japan_cg.pct_change()
Japan_cg = (1 + Japan_cg['Japan Revenue']).cumprod() - 1

Mexico['Year']=Mexico.index.year
Mexico_cg=Mexico.drop(['ProductID', 'Zip', 'Units'], axis=1)
Mexico_cg=Mexico_cg.groupby('Year').sum()
Mexico_cg.rename(columns={'Revenue': 'Mexico Revenue'}, inplace=True)
Mexico_cg=Mexico_cg.pct_change()
Mexico_cg = (1 + Mexico_cg['Mexico Revenue']).cumprod() - 1

Nigeria['Year']=Nigeria.index.year
Nigeria_cg=Nigeria.drop(['ProductID', 'Zip', 'Units'], axis=1)
Nigeria_cg=Nigeria_cg.groupby('Year').sum()
Nigeria_cg.rename(columns={'Revenue': 'Nigeria Revenue'}, inplace=True)
Nigeria_cg=Nigeria_cg.pct_change()
Nigeria_cg = (1 + Nigeria_cg['Nigeria Revenue']).cumprod() - 1

UnitedStates['Year']=UnitedStates.index.year
UnitedStates_cg=UnitedStates.drop(['ProductID', 'Zip', 'Units'], axis=1)
UnitedStates_cg=UnitedStates_cg.groupby('Year').sum()
UnitedStates_cg.rename(columns={'Revenue': 'USA Revenue'}, inplace=True)
UnitedStates_cg=UnitedStates_cg.pct_change()
UnitedStates_cg = (1 + UnitedStates_cg['USA Revenue']).cumprod() - 1

#merge 
World_cg=pd.concat([Australia_cg, Canada_cg, Germany_cg, Japan_cg, Mexico_cg, Nigeria_cg, UnitedStates_cg], join='inner', axis=1)
print(World_cg)

#plot
plt.clf()
World_cg.plot(kind='line', title='Worldwide Revenue Growth (2011 to 2018)')
plt.ylabel("Revenue")
plt.xlabel("")
plt.grid(True)
plt.show() 
#Based on the plot, Mexico, USA, and Australia have the top cumulative rev growth b/t 2011 to 2018.

#4) List the top 5 best-selling products in the world
#setup
World.dtypes
World['ProductID']=World['ProductID'].astype(str)
product.dtypes
product['ProductID']=product['ProductID'].astype(str)

#merge
World_products=pd.merge(World, product, on='ProductID')
World_products.dtypes
World_products=World_products.drop(['Zip', 'Category', 'ManufacturerID', 'Price USD', 'Revenue'], axis=1)

#finally, groupby Product and analyze results
World_products=World_products.groupby('Product').sum()
World_products=World_products.sort_values(by='Units', ascending=False).head(5)
print(World_products)

#5) List the top 5 best-selling products in Japan
Japan.dtypes
Japan['ProductID']=Japan['ProductID'].astype(str)
Japan_products=pd.merge(Japan, product, on='ProductID')
Japan_products.dtypes
Japan_products=Japan_products.drop(['Zip', 'Category', 'ManufacturerID', 'Price USD', 'Revenue', 'Year'], axis=1)
Japan_products=Japan_products.groupby('Product').sum()
Japan_products=Japan_products.sort_values(by='Units', ascending=False).head(5)
print(Japan_products)

#6) What is the most expensive product?
product_sorted=product.sort_values(by='Price USD', ascending=False).head(1)
product_sorted=product_sorted.drop(['Category', 'ManufacturerID'], axis=1)
print(product_sorted)

#7) What is the biggest [manufacturer] contributor to revenue in the world?
World.dtypes
product.dtypes
manufacturer.dtypes
manufacturer['ManufacturerID']=manufacturer['ManufacturerID'].astype(str)
World_manufacturer=pd.merge(World, product, on='ProductID')
World_manufacturer.dtypes
World_manufacturer['ManufacturerID']=World_manufacturer['ManufacturerID'].astype(str)
World_manufacturer=World_manufacturer.drop(['Units', 'Zip', 'Country', 'Category', 'Price USD', 'ProductID', 'Product'], axis=1)
World_manufacturer=pd.merge(World_manufacturer, manufacturer, on='ManufacturerID')
World_manufacturer=World_manufacturer.groupby('Manufacturer').sum()
World_manufacturer=World_manufacturer.sort_values(by='Revenue', ascending=False).head(1)
print(World_manufacturer)

#8) List the bottom 3 [manufacturer] contributor to revenue in Nigeria
Nigeria.dtypes
Nigeria['ProductID']=Nigeria['ProductID'].astype(str)
Nigeria_manufacturer=pd.merge(Nigeria, product, on='ProductID')
Nigeria_manufacturer.dtypes
Nigeria_manufacturer['ManufacturerID']=Nigeria_manufacturer['ManufacturerID'].astype(str)
Nigeria_manufacturer=Nigeria_manufacturer.drop(['Units', 'Zip', 'Country', 'Category', 'Price USD', 'ProductID', 'Product', 'Year'], axis=1)
Nigeria_manufacturer=pd.merge(Nigeria_manufacturer, manufacturer, on='ManufacturerID')
Nigeria_manufacturer=Nigeria_manufacturer.groupby('Manufacturer').sum()
Nigeria_manufacturer=Nigeria_manufacturer.sort_values('Revenue', ascending=True).head(3)
print(Nigeria_manufacturer)

#9) Provide the list of products with no revenue in 2017
#setup World df
World.dtypes
World['ProductID']=World['ProductID'].astype(int)
World_2017=World.drop(['Units'], axis=1)
World_2017=World_2017[(World_2017.index.year==2017)]
World_2017=World_2017.groupby('ProductID').sum()

#setup product df
product.dtypes
product['ProductID']=product['ProductID'].astype(int)
product_filtered=product.drop(['Category', 'ManufacturerID', 'Price USD', 'Product'], axis=1)
product_filtered=product_filtered.set_index('ProductID')

#combine dfs and calc
World_2017=pd.concat([World_2017, product_filtered])
World_2017=World_2017.sort_index(ascending=True)
World_2017=World_2017.groupby('ProductID').sum()
World_2017=World_2017[(World_2017['Revenue'] <= 0)]
print(World_2017)

#10) Provide the list of products with no revenue in 2017 & 2018
#setup World df
World_2017_2018=World.drop(['Units'], axis=1)
World_2017_2018=World_2017_2018[(World_2017_2018.index.year>=2017)]
World_2017_2018=World_2017_2018.groupby('ProductID').sum()

#combine dfs and calc
World_2017_2018=pd.concat([World_2017_2018, product_filtered])
World_2017_2018=World_2017_2018.sort_index(ascending=True)
World_2017_2018=World_2017_2018.groupby('ProductID').sum()
World_2017_2018=World_2017_2018[(World_2017_2018['Revenue'] <= 0)]
print(World_2017_2018)

#11) Provide a list of the top 3 products with the most important growth (in value) between 2016 to 2017
#prep 2016
World_2016=World.drop(['Country','Units', 'Zip'], axis=1)
World_2016.rename(columns={'Revenue': 'Revenue 2016'}, inplace=True)
World_2016=World_2016[(World_2016.index.year==2016)].groupby('ProductID').sum()

#prep 2017
World_2017_II=World.drop(['Country','Units', 'Zip'], axis=1)
World_2017_II.rename(columns={'Revenue': 'Revenue 2017'}, inplace=True)
World_2017_II=World_2017_II[(World_2017_II.index.year==2017)].groupby('ProductID').sum()

#merge 2016 and 2017 dfs and calculate growth %
World_2016.dtypes
World_2017_II.dtypes
World_growth_II=pd.merge(World_2016, World_2017_II, on='ProductID')
World_growth_II['Growth %'] = (World_growth_II['Revenue 2017'] / World_growth_II['Revenue 2016'] - 1)
World_growth_II=World_growth_II.sort_values(by='Growth %', ascending=False).head(3) #top 3 
World_growth_II.dtypes
product.dtypes
World_growth_II=pd.merge(World_growth_II, product, on='ProductID')
World_growth_II=World_growth_II.set_index('Product')
World_growth_II=World_growth_II.drop(['Category','ManufacturerID', 'Price USD'], axis=1)
print(World_growth_II)

#12) Is Blue Corporate a seasonal business? (Y/N)
#If yes: decompose this time series between seasonality & trend?

#setup
World_seasonal=World.drop(['ProductID','Zip', 'Units', 'Country'], axis=1)
World_seasonal=World_seasonal.resample('M').sum()

#setup based on Month
World.dtypes
World_seasonal_I=World.drop(['ProductID','Zip', 'Units', 'Country'], axis=1)
World_seasonal_I['Month']=World_seasonal_I.index.month
World_seasonal_I=World_seasonal_I.groupby('Month').sum()

#setup based on Day
World_seasonal_II=World.drop(['ProductID','Zip', 'Units', 'Country'], axis=1)
World_seasonal_II['Day']=World_seasonal_II.index.day
World_seasonal_II=World_seasonal_II.groupby('Day').sum()

#initial plot
plt.clf()
World_seasonal.plot(kind='line', title='Worldwide Revenue by Year (2011 to 2018)')
plt.ylabel("Revenue")
plt.xlabel("")
plt.grid(True)
plt.show() 
#Based on the initial plot, we notice a pattern. Further plots are necessary to determine trend details.

#plot based on Month
plt.clf()
World_seasonal_I.plot(kind='line', title='Worldwide Revenue by Month (2011 to 2018)')
plt.ylabel("Revenue")
plt.xlabel("")
plt.grid(True)
plt.show() 
#Based on the plot, Blue Corp. appears to be a seasonal business. Sales are strongest in Spring and weakest in Winter.

#plot based on Day
plt.clf()
World_seasonal_II.plot(kind='line', title='Worldwide Revenue by Day (2011 to 2018)')
plt.ylabel("Revenue")
plt.xlabel("")
plt.grid(True)
plt.show() 
#Based on the second plot, Blue Corp. sales spike on the 15th and 30th of the month. Not super informative.

#If yes: decompose this time series between seasonality & trend?
#decompose by revenue
World_decompose=World.drop(['ProductID','Zip', 'Units', 'Country'], axis=1)
World_decompose=World_decompose.resample('M').sum()
World_decompose=seasonal_decompose(World_decompose, model='additive', freq=12)
World_decompose.plot()

#NAM ANALYSIS
#13) Provide revenue of USA & Canada in January 2015
NorthAmerica=pd.concat([Canada, UnitedStates])
NorthAmerica_Jan_2015=NorthAmerica[(NorthAmerica.index.year==2015)]
NorthAmerica_Jan_2015=NorthAmerica_Jan_2015[(NorthAmerica_Jan_2015.index.month==1)]
NorthAmerica_Jan_2015.dtypes
NorthAmerica_Jan_2015=NorthAmerica_Jan_2015.drop(['ProductID', 'Zip', 'Units', 'Year'], axis=1)
NorthAmerica_Jan_2015=NorthAmerica_Jan_2015.groupby('Country').sum()
NorthAmerica_Jan_2015.loc["Total"] = NorthAmerica_Jan_2015.sum()
print(NorthAmerica_Jan_2015)

#14) In USA what is the top 3 states contributor to revenue?
UnitedStates.dtypes
geo.dtypes
geo_filtered=geo[(geo['Country'] == 'USA')]
UnitedStates_detail=pd.merge(UnitedStates, geo_filtered, on='Zip')
UnitedStates_detail.dtypes
UnitedStates_detail_filter=UnitedStates_detail.drop(['ProductID', 'Zip', 'Units', 'Country_x', 'Year', 'City', 'Region', 'District', 'Country_y' ], axis=1)
UnitedStates_detail_filter=UnitedStates_detail_filter.groupby('State').sum()
UnitedStates_detail_filter=UnitedStates_detail_filter.sort_values('Revenue', ascending=False).head(3)
print(UnitedStates_detail_filter)

#15) In USA what is the region with the most important share of revenue?
UnitedStates_detail_filter2=UnitedStates_detail.drop(['ProductID', 'Zip', 'Units', 'Country_x', 'Year', 'City', 'State', 'District', 'Country_y' ], axis=1)
UnitedStates_detail_filter2=UnitedStates_detail_filter2.groupby('Region').sum()
UnitedStates_detail_filter2=UnitedStates_detail_filter2.sort_values('Revenue', ascending=False).head(1)
print(UnitedStates_detail_filter2)

#16) In USA for 2014, what is the top 3 states with the most important volume of unit?
UnitedStates_detail.dtypes
UnitedStates_detail['Year']=UnitedStates_detail['Year'].astype(int)
UnitedStates_detail_filter3=UnitedStates_detail.drop(['ProductID', 'Zip', 'Revenue', 'Country_x', 'City', 'Region', 'District', 'Country_y' ], axis=1)
UnitedStates_detail_filter3=UnitedStates_detail_filter3[(UnitedStates_detail_filter3['Year'] == 2014)]
UnitedStates_detail_filter3=UnitedStates_detail_filter3.groupby('State').sum()
UnitedStates_detail_filter3=UnitedStates_detail_filter3.drop(['Year' ], axis=1)
UnitedStates_detail_filter3=UnitedStates_detail_filter3.sort_values('Units', ascending=False).head(3)
print(UnitedStates_detail_filter3)

#17) In Canada for 2017, what is the state with the most important volume of unit and revenue?
#setup
Canada.dtypes
geo.dtypes
geo_filtered=geo[(geo['Country'] == 'Canada')]
Canada_detail=pd.merge(Canada, geo_filtered, on='Zip')
Canada_detail.dtypes

#analysis of units
Canada_detail_filter=Canada_detail.drop(['ProductID', 'Zip', 'Country_x', 'City', 'Region', 'District', 'Country_y' ], axis=1)
Canada_detail_filter=Canada_detail_filter[(Canada_detail_filter['Year'] == 2017)]
Canada_detail_filter=Canada_detail_filter.groupby('State').sum()
Canada_detail_filter=Canada_detail_filter.drop(['Year' ], axis=1)
Canada_detail_filter_units=Canada_detail_filter.sort_values('Units', ascending=False).head(1)
Canada_detail_filter_units=Canada_detail_filter_units.drop(['Revenue' ], axis=1)
print(Canada_detail_filter_units) #Canadian State with highest volume of unit in 2017

#analysis of revenue
Canada_detail_filter_rev=Canada_detail_filter.sort_values('Revenue', ascending=False).head(1)
Canada_detail_filter_rev=Canada_detail_filter_rev.drop(['Units' ], axis=1)
print(Canada_detail_filter_rev) #Canadian State with highest volume of revenue in 2017

#19) WW forecast for 2019 using ARIMA Model (Simple Model)
#modules ARIMA
from statsmodels.tsa.arima_model import ARIMA
import warnings
warnings.filterwarnings('ignore')
from sklearn.metrics import mean_squared_error
from statsmodels.tsa.stattools import adfuller
import statsmodels.api as sm

adfuller = adfuller(World_seasonal['Revenue'])
print("ADF test statistic: ", str(adfuller[0]))
print("p-value: ", str(adfuller[1]))

#ARIMA
p = 1
q = 1
d = 1
World_model=sm.tsa.statespace.SARIMAX(World_seasonal['Revenue'], order=(p,q,d), seasonal_order=(1,1,1,12))
World_model_arima=World_model.fit()
World_model_arima.summary()

#forecast
World_forecast = World_model_arima.predict(start=95, end=107)
World_forecast = pd.DataFrame(World_forecast)
World_actual_forecast = pd.concat([World_seasonal,World_forecast])
World_actual_forecast = World_actual_forecast[['Revenue','predicted_mean']]
World_actual_forecast = World_actual_forecast.rename(columns={'Revenue':'Actual','predicted_mean':'Forecast'})

#plot 2011 to 2018 actual plus 2019 forecast
World_actual_forecast.plot()
plt.title('Worldwide Revenue 2011-2018 Actuals + 2019 Forecast')
plt.xlabel('Date')
plt.ylabel('Revenue')
plt.show()

#export to Excel
#determining the name of the file
file_name = 'Forecast.xlsx'
  
# saving the excel
World_actual_forecast.to_excel(file_name)
print('DataFrame is written to Excel File successfully.')