import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
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
date_time_log = SHEET.worksheet('date_time')                                     # Date Time format error log

# define the sheets to be user for outlier output
atmos_outlier_log = SHEET.worksheet('atmos_outliers')
wind_outlier_log = SHEET.worksheet('wind_outliers')
wave_outlier_log = SHEET.worksheet('wave_outliers')
temp_outlier_log = SHEET.worksheet('temp_outliers')


def df_to_list_of_lists(df):
    """
    Function to convert dataframe into list of lists.
    This is used by the outlier functions and the updater log functions
    to write the dataframes to the respective google sheet
    """
    return [df.columns.tolist()] + df.values.tolist()


# Initialize the sheets for use
def clear_all_sheets():
    """
    Function to clear all the google sheets data each time
    the app is run
    """
    print("Starting Google Sheet Initialisation\n")
    sheets = [
        validated_master_data, user_data_report, session_log, error_log,
        atmos_outlier_log, wind_outlier_log, wave_outlier_log, temp_outlier_log, date_time_log
    ]
    for sheet in sheets:
        sheet.clear()

    print("Finished Google Sheet Initialisation\n")


def load_marine_data_input_sheet():
    """
    Function to Initialise the session log and the error logs
    Then the master data dataframe is populated with the values
    of the master data google sheet
    """
    # Initialise log lists
    session_log_data = [['Session Log Started'], [str(pd.Timestamp.now())]]
    error_log_data = [['Error Log Started'], [str(pd.Timestamp.now())]]

    # write - start loading master data
    print("\nLoading Master Data          <<<<<\n")
    session_log_data.append(['Master Data Started Loading'])
    session_log_data.append([str(pd.Timestamp.now())])

    # Create Dataframe with all data from google input sheet
    print("     Master Data Loading Start")
    master_data = unvalidated_master_data.get_all_values()
    print("     Master Data Loading Complete\n")
    session_log_data.append(['Master Data Finished Loading'])
    session_log_data.append([str(pd.Timestamp.now())])
    print("Master Data Load Completed     <<<<<\n")

    # return the dataframe with the master data from the sheet
    return master_data, session_log_data, error_log_data


def check_for_outliers(df):
    """
    The purpose of the function is to check for outliers
    using Z-Score method
    """
    z_scores = zscore(df)                            # calculate a z score for all values in dataframe
    abs_z_scores = abs(z_scores)                     # take the absolute value
    outliers = (abs_z_scores > 3).any(axis=1)        # creates a bolean based on whether the absolute number is  > 3
    return df[outliers]                              # return the dataframe of the outliers identified


def validate_master_data(master_data, session_log_data, error_log_data):
    """
    This purpose of this function is to take the masterdata set and
    validate it for errors, based on:
    - Missing Values
    - Duplicate Rows
    - Outliers
    - Inconsistancy
    I create a dataframe taking input from the marine_data_m2 masterdata.
    """
    # Initialise log lists
    date_time_log_data = []  # Initialize date_time_log_data here

    print("\nData Validation Started       <<<<<\n\n\n")
    validated_data_df = pd.DataFrame()

    # Write a heading to the error log in google sheets with time/date stamp
    try:
        # Create dataframe to work with
        df = pd.DataFrame(master_data[1:], columns=master_data[0])

        # pick out the specific columns to be used in the application and create a master data frame
        master_df = df[['time', 'AtmosphericPressure', 'WindDirection',
                        'WindSpeed', 'Gust', 'WaveHeight', 'WavePeriod',
                        'MeanWaveDirection', 'AirTemperature', 'SeaTemperature', 'RelativeHumidity']]

        session_log_data.append(['Data Validation Started <<<<<<<<<<'])
        session_log_data.append([str(pd.Timestamp.now())])

        # Validate for Missing Values
        print("Validating missing values started       <<<<<\n")
        session_log_data.append(['Checking For Missing Values'])
        session_log_data.append([str(pd.Timestamp.now())])

        pd.set_option('future.no_silent_downcasting', True)
        master_df = master_df.replace(to_replace=['nan', 'NaN', ''], value=np.nan)
        missing_values = master_df.isnull().sum()
        # If there are missing values, write to the session log and error log
        if missing_values.any():
            print("     We found rows with missing values")
            print("     Please check the error log")
            session_log_data.append(['We found missing values in the master data'])
            session_log_data.append([str(pd.Timestamp.now())])
            error_log_data.append(['Missing Values       <<<<<'])
            # append missing values in a column
            for column, count in missing_values.items():
                if count > 0:
                    error_log_data.append([f"{column}: {count} missing values"])
        else:
            print("     There were no row swith missing data\n")
            session_log_data.append(['We found no missing values in the master data'])

        # Remove rows with missing values
        if missing_values.any():
            session_log_data.append(['Removing data with no values from master data'])
            session_log_data.append([str(pd.Timestamp.now())])
            missing_values_removed_df = master_df.dropna()
            missing_values = missing_values_removed_df.isnull().sum()
            print("     Rows with missing data have been removed\n")

        print("Validating missing values completed     <<<<<\n\n\n")

        # Check for duplicate rows
        print("Validating duplicates started       <<<<<\n")
        duplicates_found = missing_values_removed_df.duplicated(keep=False).sum()
        if duplicates_found:
            print(f"     There are {duplicates_found} duplicates in the working data set")
            print("     Please check the error log")
            session_log_data.append([f'Duplicates found in data set: {duplicates_found}'])
            session_log_data.append([str(pd.Timestamp.now())])
            error_log_data.append(['Duplicate Rows Found'])
            duplicates_df = missing_values_removed_df[missing_values_removed_df.duplicated(keep=False)]

            # Format duplicates_df for column-wise insertion
            duplicates_list_of_lists = duplicates_df.values.tolist()
            # Add header for duplicates (optional)
            duplicates_header = [duplicates_df.columns.tolist()]
            # Combine header and data
            formatted_duplicates = duplicates_header + duplicates_list_of_lists
            # Append to the error log data
            error_log_data.append(['Duplicate Rows Data'])
            error_log_data.extend(formatted_duplicates)
            no_duplicates_df = missing_values_removed_df.drop_duplicates(keep='first')
            print("     Duplicates have been removed\n")
        else:
            print("     No duplicates found in the working data set.\n")
            no_duplicates_df = missing_values_removed_df

        print("Validating duplicates completed     <<<<<\n\n\n")

        # Check for outliers
        print("Outlier Validation Started       <<<<<\n")
        session_log_data.append(['Outlier Validation Started'])
        session_log_data.append([str(pd.Timestamp.now())])

        numeric_df = no_duplicates_df.apply(lambda col: col.map(lambda x: pd.to_numeric(x, errors='coerce')))

        atmospheric_outliers = check_for_outliers(numeric_df[['AtmosphericPressure']])
        wind_outliers = check_for_outliers(numeric_df[['WindSpeed', 'Gust']])
        wave_outliers = check_for_outliers(numeric_df[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']])
        temp_outliers = check_for_outliers(numeric_df[['AirTemperature', 'SeaTemperature']])

        # Update Outlier Sheets
        if not atmospheric_outliers.empty:
            atmos_outlier_log.update(df_to_list_of_lists(atmospheric_outliers), 'A1')
            print("     Atmospheric Outliers Were Found: Check Atmos Outlier Log")
        if not wind_outliers.empty:
            wind_outlier_log.update(df_to_list_of_lists(wind_outliers), 'A1')
            print("     Wind Outliers Were Found:        Check Wind Outlier Log")
        if not wave_outliers.empty:
            wave_outlier_log.update(df_to_list_of_lists(wave_outliers), 'A1')
            print("     Wave Outliers Were Found:        Check Wave Outlier Log")
        if not temp_outliers.empty:
            temp_outlier_log.update(df_to_list_of_lists(temp_outliers), 'A1')
            print("     Temp Outliers Were Found:        Check Temp Outlier Log\n")

        print("Outlier Validation Completed     <<<<<\n\n\n")

        # Check for date inconsistencies
        print("Date Validation Started       <<<<<\n")
        session_log_data.append(['Date and Time Validation Started'])
        session_log_data.append([str(pd.Timestamp.now())])
        date_time_pattern = r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
        inconsistent_date_format = no_duplicates_df[~no_duplicates_df['time'].str.match(date_time_pattern)]

        if inconsistent_date_format.empty:
            print("     No Incorrect Date Formats Found")
            validated_data_df = no_duplicates_df.copy()
            validated_data_df['time'] = pd.to_datetime(validated_data_df['time'], format='%Y-%m-%dT%H:%M:%SZ')
            validated_data_df['time'] = validated_data_df['time'].dt.strftime('%d-%m-%YT%H:%M:%S')
        else:
            print("     Incorrect Date Formats Found\n")
            date_time_log_data.append(['Inconsistent Date and Time Formats'])
            date_time_log_data.append(inconsistent_date_format['time'].fillna('').astype(str).values.tolist())
            validated_data_df = no_duplicates_df[no_duplicates_df['time'].str.match(date_time_pattern)].copy()
            validated_data_df['time'] = pd.to_datetime(validated_data_df['time'], format='%Y-%m-%dT%H:%M:%SZ')
            validated_data_df['time'] = validated_data_df['time'].dt.strftime('%d-%m-%YT%H:%M:%S')

        print("     Incorrect Date Formats Removed\n")
        print("Date Validation Completed     <<<<<\n")
        session_log_data.append(['Data Validation Ended <<<<<<<<<<'])
        session_log_data.append([str(pd.Timestamp.now())])

        # Convert all elements in logs to strings
        session_log_data = [[str(item) for item in sublist] for sublist in session_log_data]
        error_log_data = [[str(item) for item in sublist] for sublist in error_log_data]
        date_time_log_data = [[str(item) for item in sublist] for sublist in date_time_log_data]

        # Write all accumulated data at once
        # session_log.update(session_log_data, 'A10')
        # error_log.update(error_log_data, 'A10')
        date_time_log.update(date_time_log_data, 'A1')  # Log inconsistent date formats here

    except Exception as e:
        session_log_data.append([f"An error occurred during validation: {e}"])
        print(f"An error occurred during validation: {e}")

    finally:
        # Convert log lists to strings and update Google Sheets
        session_log.update(df_to_list_of_lists(pd.DataFrame(session_log_data)), 'A1')
        error_log.update(df_to_list_of_lists(pd.DataFrame(error_log_data)), 'A1')
        date_time_log.update(df_to_list_of_lists(pd.DataFrame(date_time_log_data)), 'A1')

    print("\nData Validation Completed     <<<<<\n\n\n")
    print("Writing Validated Data To Google Sheets Started      <<<<<\n")
    set_with_dataframe(validated_master_data, validated_data_df)
    print("Writing Validated Data To Google Sheets Completed    <<<<<\n")

    return validated_data_df


def get_user_dates(validated_df):
    """
    This function:
    - displays the date range available in master data
    - asks for input start date
    - asks for input end date
    """
    # Convert the time column in validated_df to date time
    validated_df['time'] = pd.to_datetime(validated_df['time'], format='%d-%m-%YT%H:%M:%S')

    # Get the first and last dates available in the data frame
    df_first_date = validated_df['time'].min().strftime('%d-%m-%Y')
    df_last_date = validated_df['time'].max().strftime('%d-%m-%Y')
    print(f"Available data range: {df_first_date} to {df_last_date}")

    def validate_input_dates(date_str, reference):
        """
        function to take input values for start date and end date
        checking to see they are within the ranges available
        """
        try:
            # Convert the input to a datetime object
            date = pd.to_datetime(date_str, format='%d-%m-%Y')

            # Extract day, month, and year to perform further checks
            day, month, year = date_str.split('-')
            day = int(day)
            month = int(month)
            year = int(year)

            # check to see that month is valid 1 - 12
            if month < 1 or month > 12:
                raise ValueError("Month must be between 1 and 12.")

            # Check the number of days in the month within range 1 -31
            if day < 1 or day > 31:
                raise ValueError("Day must be between 1 and 31.")

            # Check that months, 4,6,9 and 11 are betwen 1 and 30
            if month in [4, 6, 9, 11] and day > 30:
                raise ValueError(f"Month {month} only has 30 days.")

            # Allow for leap year input
            if month == 2:
                if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                    if day > 29:
                        raise ValueError("February in a leap year has only 29 days.")
                else:
                    if day > 28:
                        raise ValueError("February has only 28 days in a non-leap year.")

            # Ensure the date falls within the available range
            if date < pd.to_datetime(df_first_date) or date > pd.to_datetime(df_last_date):
                raise ValueError(f"Date {date_str} is outside the available data range.")
            return date
        except ValueError as ve:
            # Re-raise the error with a custom message
            raise ValueError(f"Invalid date: {ve}. Please enter the date in 'dd-mm-yyyy' format.")

    # Prompt user for start date, dont go to end date until date is acceptable and formatted
    while True:
        user_input_start_date = input(f"Please enter the start date\n (in the format 'dd-mm-yyyy'\nwithin {df_first_date} and {df_last_date}): ")
        try:
            user_input_start_date = validate_input_dates(user_input_start_date, "start")
            break  # Exit loop if the date is valid
        except ValueError as e:
            print(e)  # Print error message and prompt again

    # Prompt user for end date
    while True:
        user_input_end_date = input(f"Please enter the end date\n (in the format 'dd-mm-yyyy'\n within {df_first_date} and {df_last_date}): ")
        try:
            user_input_end_date = validate_input_dates(user_input_end_date, "end")
            if user_input_end_date < user_input_start_date:
                raise ValueError("The end date cannot be earlier than the start date.")
            break  # Exit loop if the date is valid
        except ValueError as e:
            print(e)  # Print error message and prompt again
        


    # Convert validated start and end dates to string format for further processing
    user_input_start_date_str = user_input_start_date.strftime('%d-%m-%Y')
    user_input_end_date_str = user_input_end_date.strftime('%d-%m-%Y')


    return user_input_start_date_str, user_input_end_date_str


# Initialise the validated dataframe
validated_df = None


def data_initialisation_and_validation():
    """
    This function:
    - Clears all the sheets
    - loads the marine data for validation
    - Validates the data
    - Stores the validated data for use in the session
    Its only run once per session
    """
    global validated_df
    clear_all_sheets()                                                                    # Initialise Sheets On Load
    master_data, session_log_data, error_log_data = load_marine_data_input_sheet()        # Load the marine date for validation
    validated_df = validate_master_data(master_data, session_log_data, error_log_data)    # Create a validated data frame for use in the app
    print("Validated DF\n\n\n")
    print(validated_df)

def main():
    """
    Main function that controls all the functionality of the application
    - Loading Master Data
    - Data validation
    - Getting user input for date ranges
    - Filtering the dataframe based on date ranges from user
    """
    global validated_df  # status of validated_df

    # Check to see if the data has been validated
    if validated_df is None:
        print("Error: Data has not been validated")
        return  # Exit the function early if validation is not done


    # convert the time columnin df to just date for comparison drop h:m:s
    validated_df['time'] = pd.to_datetime(validated_df['time'], format='%d-%m-%YT%H:%M:%S', errors='coerce', dayfirst=True)   
    # Extract the date component only, ignoring the time
    validated_df['date_only'] = validated_df['time'].dt.strftime('%d-%m-%Y')
    
    print(validated_df['date_only'])


    # Get dates from user to interrogate validated data frame
    user_input_start_date_str, user_input_end_date_str = get_user_dates(validated_df)
    print(validated_df['time'])


    # Convert user input dates to the format used in validated_df
    user_input_start_date = user_input_start_date_str
    user_input_end_date = user_input_end_date_str
    print(f"Start Date: {user_input_start_date}")
    print(f"Start Date: {user_input_end_date}")


    # Filter the dataframe based on the date range
    date_filtered_df = validated_df[
        (pd.to_datetime(validated_df['date_only'], format='%d-%m-%Y') >= user_input_start_date) &
        (pd.to_datetime(validated_df['date_only'], format='%d-%m-%Y') <= user_input_end_date)
    ]
    print("Date Filter For Data Selection:")
    print(date_filtered_df)

    # Prepare date_filtered_df for use as output
    # Change the date format from yyyy-mm-dd 00:00:00 to dd-mm-yyyy 00:00:00
    # date_filtered_df['time'] = pd.to_datetime(date_filtered_df['time']).dt.strftime(('%d-%m-%Y %H:%M:%S'))
    # date_filtered_df.loc[:, 'time'] = pd.to_datetime(date_filtered_df['time']).dt.strftime('%d-%m-%Y %H:%M:%S')
    # print(date_filtered_df)

    # Create a copy of date_filtered_df to avoid modifying the original DataFrame
    data_manipulation_df = date_filtered_df.copy()

    # Convert the 'time' column to the desired format in the copied DataFrame
    fdata_manipulation_df['time'] = pd.to_datetime(data_manipulation_df['time']).dt.strftime('%d-%m-%Y %H:%M:%S')

    # Now formatted_df will have the 'time' column formatted as 'dd-mm-yyyy 00:00:00'
    print(formatted_df)

    # Output Selection Options
    print("\nSelect the data you want to display:")
    print("1: All Data")
    print("2: Atmospheric Pressure")
    print("3: Wind Speed and Gust")
    print("4: Wave Height, Wave Period, and Mean Wave Direction")
    print("5: Air Temperature and Sea Temperature")

    # take users selected number
    selection = input("Enter the number corresponding to your selection")

    # Create relevant data subset dataframes for processing
    if selection == 1:
        selected_columns = ['time', 'AtmosphericPressure', 'WindDirection',
                            'WindSpeed', 'Gust', 'WaveHeight', 'WavePeriod',
                            'MeanWaveDirection', 'AirTemperature', 'SeaTemperature', 'RelativeHumidity']
    elif selection == 2:
        selected_columns = ['time', 'AtmosphericPressure']
    elif selection == 3:
        selected_columns = ['time', 'WindSpeed', 'Gust']
    elif selection == 4:
        selected_columns = ['time', 'WaveHeight', 'WavePeriod', 'MeanWaveDirection']
    elif selection == 5:
        selected_columns = ['time', 'AirTemperature', 'SeaTemperature']

    # Print selected columns for debugging
    print(f"Selected Columns: {selected_columns}")   


# Initialise the sheets and validate the data
data_initialisation_and_validation()


main()
