import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials
from scipy.stats import zscore

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
validated_master_data = SHEET.worksheet('validated_master_data')                 # Validated Master Data for use
user_data_report = SHEET.worksheet('user_data_report')                           # Sheet containing requested data output
data_error_log = SHEET.worksheet('gael_force_error_log')                         # Errors sent to log


def load_marine_data_input_sheet():
    print("Loading Master Data\n")
    master_data = unvalidated_master_data.get_all_values()                       # assign all values in master data to variable for use
    print("Master Data Load Completed\n")
    return master_data


def check_for_outliers(df):
    """
    The purpose of the function is to check for outliers
    using Z-Score method
    """
    z_scores = zscore(df)
    abs_z_scores = abs(z_scores)
    outliers = (abs_z_scores > 3).any(axis=1)
    return df[outliers]


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
    # Create dataframe to work with
    df = pd.DataFrame(master_data[1:], columns=master_data[0])

    #
    #
    # Check for missing values
    #
    #

    # pick out the specific columns to be used in the application
    master_df = df[['time', 'AtmosphericPressure', 'WindDirection',
                    'WindSpeed', 'Gust', 'WaveHeight', 'WavePeriod',
                    'MeanWaveDirection', 'Hmax', 'AirTemperature', 'SeaTemperature', 'RelativeHumidity']]

    print("Starting Data Validation\n")
    print("Checking for missing values in data set\n")
    # Validate for Missing Values
    pd.set_option('future.no_silent_downcasting', True)                           # Neccessary to include to aboid downcasting message
    master_df = master_df.replace(to_replace=['nan', 'NaN', ''], value=np.nan)    # Replace versions of nan in df with numpy nan

    # Get a count of cell values with missing data
    count_of_values_with_nan = master_df.isnull().sum()                           # get sum of missing values
    if count_of_values_with_nan.any():
        print("We found the following errors\n")                                  # If there are any
        print(count_of_values_with_nan)
        print("\n")                                                               # Print them to the screen
    else:
        print("There were no cells with missing data")                           # Print to screen message no errors
        print(count_of_values_with_nan)
        print("\n")

    # Remove any rows that have no values
    print("Starting to clean data frame\n")
    values__validated_df = master_df.dropna(subset=['time', 'AtmosphericPressure', 'WindDirection',
                                                    'WindSpeed', 'Gust', 'WaveHeight', 'WavePeriod',
                                                    'MeanWaveDirection', 'AirTemperature',
                                                    'SeaTemperature', 'RelativeHumidity'])
    count_of_nan_after_cleaning = values__validated_df.isnull().sum()
    print("The Result is")
    print(count_of_nan_after_cleaning)
    print("Missing value validation completed\n")

    #
    #
    # Check dor duplicate rows
    #
    #
    print("Validating duplicates started\n")
    count_of_duplicates = values__validated_df.duplicated().sum()
    print(f"There are {count_of_duplicates} duplicates in the working data set\n")
    duplicates_validated_df = values__validated_df.drop_duplicates(keep='first')
    print("Validating duplicates ended\n")

    #
    #
    # Check for outliers
    #
    #

    # Convert cell values to integers for mathamtical use
    df_numeric = duplicates_validated_df.apply(lambda col: col.map(lambda x: pd.to_numeric(x, errors='coerce')))

    # Create related dataframes for data comparison purposes
    atmospheric_specific_df = df_numeric[['AtmosphericPressure']]
    wind_specific_df = df_numeric[['WindSpeed', 'Gust']]
    wave_specific_df = df_numeric[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']]
    temp_specific_df = df_numeric[['AirTemperature', 'SeaTemperature']]
    print(atmospheric_specific_df)
    print(wind_specific_df)
    print(wave_specific_df)
    print(temp_specific_df)

    # Create outliers from specific data frames
    atmospheric_outliers = check_for_outliers(atmospheric_specific_df)
    wind_outliers = check_for_outliers(wind_specific_df)
    wave_outliers = check_for_outliers(wave_specific_df)
    temp_outliers = check_for_outliers(temp_specific_df)
    print("Atmospheric Outliers:\n", atmospheric_outliers)
    print("Wind Outliers:\n", wind_outliers)
    print("Wave Outliers:\n", wave_outliers)
    print("Temperature Outliers:\n", temp_outliers)

    #
    #
    #
    #
    #

    print("End of data validation\n")
    
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