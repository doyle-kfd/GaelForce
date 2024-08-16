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
session_log = SHEET.worksheet('session_log')                                     # Errors sent to log
error_log = SHEET.worksheet('gael_force_error_log')                              # Errors sent to log



# Initialise the sheets for use
session_log.clear()
error_log.clear()                                                              # Clear the error log for output
user_data_report.clear()                                                         # clear the User data report sheet for output


def load_marine_data_input_sheet():
    # Open Session log for recording output
    # write session headers to session log and error log
    session_log.update([['Session Log Started']], 'A1')
    error_log.update([['Error Log Started']], 'A1')
    log_timestamp = pd.Timestamp.now()
    session_log.update([['{}'.format(log_timestamp)]], 'F1')
    error_log.update([['{}'.format(log_timestamp)]], 'F1')

    # write - start loading master data
    print("Loading Master Data\n")
    session_log.update([['Master Data Started Loading']], 'A3')
    log_timestamp = pd.Timestamp.now()
    session_log.update([['{}'.format(log_timestamp)]], 'F3')
    master_data = unvalidated_master_data.get_all_values()                       # assign all values in master data to variable for use
    session_log.update([['Master Data Finished Loading']], 'A4')
    log_timestamp = pd.Timestamp.now()
    session_log.update([['{}'.format(log_timestamp)]], 'F4')
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
    # Write a heading to the error log in google sheets with time/date stamp
    try:
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
                        'MeanWaveDirection', 'AirTemperature', 'SeaTemperature', 'RelativeHumidity']]

        print("Starting Data Validation\n")                                           # Starting validation
        session_log.update([['Starting Data Validation']], 'A6')
        log_timestamp = pd.Timestamp.now()
        session_log.update([['{}'.format(log_timestamp)]], 'F6')

        print("Checking for missing values in data set\n")
        session_log.update([['Checking For Missing Values']], 'A8')
        log_timestamp = pd.Timestamp.now()
        session_log.update([['{}'.format(log_timestamp)]], 'F8')
        # Validate for Missing Values
        pd.set_option('future.no_silent_downcasting', True)                           # Neccessary to include to aboid downcasting message
        master_df = master_df.replace(to_replace=['nan', 'NaN', ''], value=np.nan)    # Replace versions of nan in df with numpy nan

        # Get a count of cell values with missing data
        missing_values = master_df.isnull().sum()                           # get sum of missing values
        if missing_values.any():
            print("We found the following columns with missing values\n")             # If there are any
            log_timestamp = pd.Timestamp.now()
            
            session_log.update([['We found missing values in the master data, please check the error log']], 'A10')
            session_log.update([['{}'.format(log_timestamp)]], 'F10')
            error_log.update([['We found missing values in the master data, in the following columns:']], 'A3')
            error_log.update([['{}'.format(log_timestamp)]], 'F3')
            print(missing_values)
            error_log.update([['{}'.format(missing_values)]], 'A4')

            # Convert the count_of_values_with_nan Series to a DataFrame for formatting to sheet
            missing_values_df = missing_values[missing_values > 0].reset_index()
            missing_values_df.columns = ['Column', 'Missing Values']
            missing_values.columns = ['Column', 'Missing Values']

            # Write the DataFrame to the error log
            for i, row in missing_values_df.iterrows():
                error_log.update([[row['Column'], row['Missing Values']]], f'A{i+4}') 
            print("\n")                                                               # Print them to the screen
        else:
            print("There were no cells with missing data")                            # Print to screen message no errors
            print(missing_values)
            session_log.update([['We found no missing values in the master data']], 'A10')
            print("\n")

        # Remove any rows that have no values
        if missing_values.any():
            session_log.update([['Removing data with no values from master data']], 'A12')
            log_timestamp = pd.Timestamp.now()
            session_log.update([['{}'.format(log_timestamp)]], 'f12')
            print("Starting to clean data frame\n")
            values_validated_df = master_df.dropna()                                                 # Create a new df having dropped NaN values
            session_log.update([['Finished removing data with no values from master data']], 'A13')
            log_timestamp = pd.Timestamp.now()
            session_log.update([['{}'.format(log_timestamp)]], 'F13')
            count_of_nan_after_cleaning = values_validated_df.isnull().sum()
            print("The Result is")
            print(count_of_nan_after_cleaning)
            print("Missing value validation completed\n")
            print("\n\n\n")
            print(master_df)
            print("\n\n\n")
            print(values_validated_df)
            print("\n\n\n")
        #
        #
        # Check dor duplicate rows
        #
        #
        print("Validating duplicates started\n")
        count_of_duplicates = values_validated_df.duplicated().sum()
        print(f"There are {count_of_duplicates} duplicates in the working data set\n")
        duplicates_validated_df = values_validated_df.drop_duplicates(keep='first')
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
        # Check for data inconsistancy in time column
        #
        #

        # Create a pattern to search for in the time field yyyy-mm-ddT00:00:00Z


        """
        ------------------->>>>>>>>> Need to come back to this an implement more robust solution.
        ------------------->>>>>>>>> It works to show
        """


        pattern = r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'


        inconsistent_date_format = duplicates_validated_df[~duplicates_validated_df['time'].str.match(pattern)]
        """
        if  inconsistent_date_format.empty:
            print("All time values are in the correct format")
            duplicates_validated_df['time'] = pd.to_datetime(duplicates_validated_df['time'], format='%Y-%m-%dT%H:%M:%SZ')
            duplicates_validated_df['time'] = duplicates_validated_df['time'].dt.strftime('%d-%m-%y')
        else:
            print("Inconsistant time format found:\n")
            print(inconsistent_date_format['time'])
            print("\n")

        """

        # print(values_validated_df)
        print("\n\n\n\n\n")
        print(duplicates_validated_df['time'])
        print("\n\n\n\n\n")

        print("End of data validation\n")

        return duplicates_validated_df

        

    except Exception as e:
        print(f"An error occurred during validation: {e}")


def get_user_dates(validated_df):
    """
    This function:
    - displays the date range available in master data
    - asks for input start date
    - asks for input end date    
    """
    start_date = validated_df['time'].iloc[0]
    print(f"Start Date Is: {start_date}")
    print(validated_df)
    print("The dates available are:")


def main():
    master_data = load_marine_data_input_sheet()
    validated_df = validate_master_data(master_data)
    # user_input_dates = get_user_dates(validated_df)
    # print(f"User Dates Provided: {user_input_dates}")


main()
