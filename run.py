import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]


# Set up authorisation for access to Google Sheet
CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('marine_data_m2')

# define the variables used to acccess each of the sheets in the Google Sheet
unvalidated_master_data = SHEET.worksheet('marine_data_master_data_2020_2024')   # Master Data file
validated_master_data = SHEET.worksheet('validated_master_data') # Validated Master Data for use
user_data_report = SHEET.worksheet('user_data_report')           # Sheet containing requested data output
data_error_log = SHEET.worksheet('gael_force_error_log')         # Errors sent to log 

def load_marine_data_input_sheet():
    print("Loading Master Data")
    master_data = unvalidated_master_data.get_all_values()             # assign all values in master data to variable for use
    print("Master Data Load Completed")
    return master_data


def validate_master_data(master_data):
    """
    This purpose of this function is to take the masterdata set and
    validate it for errors, based on:
    - Missing Values
    - Duplicate Rows
    - Outliers
    - Inconsistancy
    I create a dataframe taking input from the marine_data_m2 masterdata.
    """
    print("Starting creation of dataframe")
    # Create dataframe to work with
    df = pd.DataFrame(master_data[1:], columns=master_data[0])
    df_focused = df[['station_id', 'longitude', 'latitude']]                    # <------------------finished here. creating new df with only columns we want
    print(df_focused)
    print("finished creating data frame")
    

    # Validate for Missing Values
    pd.set_option('future.no_silent_downcasting', True)                         # Neccessary to include to aboid downcasting message
    df = df.replace(to_replace=['nan', 'NaN', ''], value=np.nan)                # Replace versions of nan in df with numpy nan

    # Get a count of cell values with missing data
    count_of_values_with_nan = df.isnull().sum()                                         # get sum of missing values
    if count_of_values_with_nan.any():
        print("We found the following errors\n")                                # If there are any
        print(count_of_values_with_nan)                                                  # Print them to the screen
    else:
        print("There were no cells with missing data")                          # Print to screen message no errors
        print(count_of_values_with_nan)
    

    # Remove any rows that have no values
    print("Starting to clean data frame")
    new_df = df.dropna(subset=['station_id', 'longitude', 'latitude',    \
    'time', 'AtmosphericPressure', 'WindDirection', 'WindSpeed', 'Gust', \
    'WaveHeight', 'WavePeriod', 'MeanWaveDirection', 'AirTemperature',   \
    'SeaTemperature', 'RelativeHumidity'])
    print(new_df.columns)
    count_of_nan_after_cleaning = new_df.isnull().sum()
    print("The Result is")
    print(count_of_nan_after_cleaning)
    



 


    
    


    
    
    """
 

    # check for outliers
    df = df.drop(columns=(['station_id', 'longitude', 'latitude', 'time', 'QC_Flag']))
    print("Atmospheric Pressure Data")
    print(df.describe()[['AtmosphericPressure']])
    
    print("Wind Related Statistics")
    print(df.describe()[['WindDirection', 'WindSpeed', 'Gust']])
    print(cleaned_data.describe()[['WindDirection', 'WindSpeed', 'Gust']])
    print("Wave Related Statistics")
    print(df.describe()[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']])
    print(cleaned_data.describe()[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']])
    print("Temp Related Statistics")
    print(df.describe()[['AirTemperature', 'DewPoint', 'SeaTemperature', 'RelativeHumidity']])
    """





    
    


def main():
    master_data = load_marine_data_input_sheet()
    validate_master_data(master_data)



main()