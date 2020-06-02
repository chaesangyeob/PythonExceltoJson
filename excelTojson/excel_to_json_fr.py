import pandas as pd
import numpy as np
from datetime import datetime
import time
import json

# Get Market Data
df_markets_france_annual = pd.read_excel('france/Vola-Tool_1_2_France.xlsx', sheet_name='29 Year Chart',
                    header=[3], usecols='A:C')
df_markets_france_data = pd.read_excel('france/Vola-Tool_1_2_France.xlsx', sheet_name='29 Year Chart',
                    header=[13], usecols='A:C')

thirty_year_annualised_return=round(df_markets_france_annual.iloc[1,2]*100,2)
five_year_annualised_return=round(df_markets_france_annual.iloc[4,2]*100,2)

points_country = []
for row in df_markets_france_data.itertuples(index=False):
    a = row.Date.to_pydatetime()
    percent_country = row._2
    datas_country = {"year":a.year,"month":a.month,"day":a.day,"percent" : percent_country}
    points_country.append(datas_country)

markets_data = {
    "markt":{
        "fr":{
            "annualised":{
                "value": thirty_year_annualised_return,
                "value2": five_year_annualised_return
            },
            "points":points_country
        }
    }
}

# Get consequence Data
df_consequence = pd.read_excel('france/Vola-Tool_1_2_France.xlsx', sheet_name='Unpredictable Returns', index_col='FRANCE CAC ALL-TRADABLE',
                    header=[16], usecols='B:V')

country_highest = []
country_lowest = []
index = 0
    
for label,content in df_consequence.iloc[[0,1,2,3]].items():
    index +=1
    highest_percent=round(content[0]*100,2)
    highest_period= content[1].split('to')
    highest_data = {
        "percent":highest_percent,
        "year":index,
        "from":{
            "month":datetime.strptime(highest_period[0].strip(),'%b %Y').month, 
            "year":datetime.strptime(highest_period[0].strip(),'%b %Y').year}, 
        "to":{
            "month":datetime.strptime(highest_period[1].strip(),'%b %Y').month,
            "year":datetime.strptime(highest_period[1].strip(),'%b %Y').year},
        }
    
    lowest_percent=round(content[2]*100,2)
    lowest_period=content[3].split('to')
    lowest_data =  {
        "percent":lowest_percent,
        "year":index,
        "from":{
            "month":datetime.strptime(lowest_period[0].strip(),'%b %Y').month, 
            "year":datetime.strptime(lowest_period[0].strip(),'%b %Y').year}, 
        "to":{
            "month":datetime.strptime(lowest_period[1].strip(),'%b %Y').month,
            "year":datetime.strptime(lowest_period[1].strip(),'%b %Y').year},
        }

    country_highest.append(highest_data)
    country_lowest.append(lowest_data)

consequence_data ={
    "markt":{
        "fr":{
            "charts":{
                "highest":country_highest,
                'lowest':country_lowest
            }
        }
    }
}
# make json file

with open('france/markets.json','w') as market:
    json.dump(markets_data,market)

with open('france/consequence.json','w') as consequence:
    json.dump(consequence_data,consequence)
    