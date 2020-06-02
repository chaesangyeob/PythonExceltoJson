how to use excel_to_json

1---
    Place the Vola excel file where you want to place and change df, df_annual, df_consequence,df_timepoints_country,df_timepoints_europa,df_timepoints_world in excel_to_json.py

2---
    According to specific country code, change "markets_data,consequence_data,timepoints_data.
    example) liechtenstein's country code is "ch" so place "ch"
    
    timepoints_data = {
    "markt":{
        "ch":{
            "percents":timepoints_country_percents
        },
        "europa":{
            "percents":timepoints_europa_percents
        },
        "world":{
            "percents":timepoints_world_percents
        }
    }
}


3---
    run excel_to_json.py , then result json file will place where "with open(....)".



4---
    since content.json is a str. you should do hardcoding