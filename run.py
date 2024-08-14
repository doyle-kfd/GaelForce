import gspread
import pandas as pd
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
    - Nan
    I create a dataframe taking input from the marine_data_m2 masterdata.
    """
    print("Creating Dataframe")
    df = pd.DataFrame(master_data[1:], columns=master_data[0])
 
    # check for missing values in master data
    missing_values = df.isnull().sum()
    print("Missing Values")
    print(missing_values)

    # check for outliers
    df = df.drop(columns=(['station_id', 'longitude', 'latitude', 'time', 'QC_Flag']))
    print("Atmospheric Pressure Data")
    print(df.describe()[['AtmosphericPressure']])
    print("Wind Related Statistics")
    print(df.describe()[['WindDirection', 'WindSpeed', 'Gust']])
    print("Wave Related Statistics")
    print(df.describe()[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']])
    print("Temp Related Statistics")
    print(df.describe()[['AirTemperature', 'DewPoint', 'SeaTemperature', 'RelativeHumidity']])
    
    


def main():
    master_data = load_marine_data_input_sheet()
    validate_master_data(master_data)



main()