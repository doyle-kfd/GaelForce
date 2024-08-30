import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
from scipy.stats import zscore
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.exceptions import GoogleAuthError

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
unvalidated_master_data = SHEET.worksheet(
    'marine_data_master_data_2020_2024')  # Master Data file
validated_master_data = SHEET.worksheet(
    'validated_master_data')  # Validated Master Data for use
user_data_output = SHEET.worksheet(
    'user_data_output')  # Sheet containing requested data output
session_log = SHEET.worksheet('session_log')  # Errors sent to log
error_log = SHEET.worksheet('gael_force_error_log')  # Errors sent to log
date_time_log = SHEET.worksheet('date_time')  # Date Time format error log
graphical_output_sheet = SHEET.worksheet('graphical_output_data')

# define the sheets to be user for outlier output
atmos_outlier_log = SHEET.worksheet('atmos_outliers')
wind_outlier_log = SHEET.worksheet('wind_outliers')
wave_outlier_log = SHEET.worksheet('wave_outliers')
temp_outlier_log = SHEET.worksheet('temp_outliers')

# define url links to each worksheet tab
google_worksheet = (
    "Worksheet: marine_data_m2 - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing"
)

marine_master_data_url = (
    "Tab: marine_data_master_data_2020_2024 - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing0"
)

validated_master_data_ulr = (
    "Tab: validated_master_data - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing178716219"
)

user_data_output_url = (
    "Tab: user_data_output - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing911199205"
)

session_log_url = (
    "Tab: session_log - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing410413428"
)

gael_force_error_log_url = (
    "Tab: gael_force_error_log - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing1094726921"
)

atmos_outliers_url = (
    "Tab: atmos_outliers - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing410485236"
)

wind_outliers_url = (
    "Tab: wind_outliers - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing594709402"
)

wave_outliers_url = (
    "Tab: wave_outliers - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing691663590"
)

temp_outliers_url = (
    "Tab: temp_outliers - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing697757766"
)

date_time_url = (
    "Tab: date_time - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing2036646154"
)

graphical_output_data_url = (
    "Tab: graphical_output_data - URL: "
    "https://docs.google.com/spreadsheets/d/"
    "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/edit?usp=sharing2079660690"
)


def df_to_list_of_lists(df):
    """
    Function to convert dataframe into list of lists.
    This is used by the outlier functions and the updater log functions
    to write the dataframes to the respective google sheet
    """
    return [df.columns.tolist()] + df.values.tolist()

def check_google_sheet_access(credentials_path, sheet_name):
    """
    Checks the accessibility of a Google Sheet using provided credentials.

    Args:
    credentials_path (str): Path to the service account credentials JSON file.
    sheet_name (str): The name of the Google Sheet to access.

    Returns:
    gspread.models.Worksheet: The first worksheet object if successful, None otherwise.
    """
    # Define the scope
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    try:
        # Load credentials
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        client = gspread.authorize(creds)
        
        # Access the Google Sheet
        sheet = client.open(sheet_name).sheet1  # Access the first sheet
        print("Successfully accessed the Google Sheet!")
        return sheet

    except FileNotFoundError as e:
        print(f"Error: The specified credentials file was not found.\nDetails: {e}")
    except PermissionError as e:
        print(f"Error: Permission denied while accessing the credentials file.\nDetails: {e}")
    except gspread.exceptions.SpreadsheetNotFound as e:
        print(f"Error: The specified Google Sheet was not found.\nDetails: {e}")
    except gspread.exceptions.APIError as e:
        print(f"Error: There was an API error when trying to access the Google Sheet.\nDetails: {e}")
    except IOError as e:
        print(f"Error: An I/O error occurred.\nDetails: {e}")
    except Exception as e:
        print(f"An unexpected error occurred.\nDetails: {e}")
    
    return None

# Initialize the sheets for use
def initialise_google_sheets():
    """
    Function to clear all the google sheets data each time
    the app is run
    """
    print("\n")
    print("#######################################")
    print("Initialising Necessary Files To Run App")
    print("#######################################")
    print("\n")
    sheet = check_google_sheet_access('creds.json', 'marine_data_m2')
    print("Starting Google Sheet Initialisation\n")
    print(f"The link to the google sheet is:\n\n{google_worksheet}\n\n")
    sheets = [
        validated_master_data, user_data_output, session_log, error_log,
        atmos_outlier_log, wind_outlier_log, wave_outlier_log,
        temp_outlier_log, date_time_log, graphical_output_sheet
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
    print("Opening a session log.\n")
    print(f"The link to the session log is:\n\n{session_log_url}\n\n")
    print("Opening a error log.\n")
    print(f"The link to the error log is:\n\n{gael_force_error_log_url}\n\n")
    # Initialise log lists
    session_log_data = [['Session Log Started'], [str(pd.Timestamp.now())]]
    error_log_data = [['Error Log Started'], [str(pd.Timestamp.now())]]

    # write - start loading master data
    print("\nLoading Master Data          <<<<<\n")
    session_log_data.append(['Master Data Started Loading'])
    session_log_data.append([str(pd.Timestamp.now())])

    # Create Dataframe with all data from google input sheet
    print("     Master Data Loading Start")
    print
    master_data = unvalidated_master_data.get_all_values()
    print("     Master Data Loading Complete\n")
    session_log_data.append(['Master Data Finished Loading'])
    session_log_data.append([str(pd.Timestamp.now())])
    print("Master Data Load Completed     <<<<<\n")

    # return the dataframe with the master data from the sheet
    return master_data, session_log_data, error_log_data


def check_for_outliers(df):
    """
    Identifies outliers in a given DataFrame using the Z-Score method.

    The function calculates the Z-Score for each value in the DataFrame and identifies
    outliers as any value with an absolute Z-Score greater than 3. These outliers are
    typically considered to be statistically significant deviations from the mean.

    Args:
    - df (pandas.DataFrame): The input DataFrame containing numerical data to check for outliers.
    """
    z_scores = zscore(df)  # calculate a z score for all values in dataframe
    abs_z_scores = abs(z_scores)  # take the absolute value
    outliers = (abs_z_scores > 3).any(
        axis=1)  # creates a bolean based on whether the absolute number is  > 3
    return df[outliers]  # return the dataframe of the outliers identified


def validate_master_data(master_data, session_log_data, error_log_data):
    """
    Validates the master data for a marine data set by performing the 
    following checks:

    1. **Missing Values:** Identifies and logs rows with missing values.
       - Missing values are replaced with NaN for consistent handling.
       - Rows containing missing data are removed after logging the issue.

    2. **Duplicate Rows:** Detects and logs duplicate rows in the data set.
       - Duplicates are identified and removed, and the details of these 
         rows are logged.

    3. **Outliers:** Checks for outliers in specific numerical columns.
       - Outliers are detected in columns such as AtmosphericPressure, 
         WindSpeed, WaveHeight, and temperatures.
       - Outliers are logged in separate logs specific to each type of data 
         (e.g., Atmospheric, Wind, Wave, Temp).

    4. **Date Inconsistencies:** Validates the format of the date and time 
         column.
       - Ensures that date and time values match a specified pattern 
        (ISO 8601 format).
       - Logs any rows with inconsistent date formats and attempts 
         to correct them.

    The function also updates session and error logs in Google Sheets for 
    tracking the validation process.

    Args:
        master_data (list): The input master data set where the first row 
        contains column headers.
        session_log_data (list): A list to accumulate session log entries.
        error_log_data (list): A list to accumulate error log entries.
    """
    # Initialise log lists
    date_time_log_data = []  # Initialize date_time_log_data here
    print("\n\n >>>>> Validate Master Data <<<<<\n\n\n")
    print("\nData Validation Started       <<<<<\n\n\n")
    validated_data_df = pd.DataFrame()

    # Write a heading to the error log in google sheets with time/date stamp
    try:
        # Create dataframe to work with
        df = pd.DataFrame(master_data[1:], columns=master_data[0])

        # pick out the specific columns to be used in the application and 
        # create a master data frame
        master_df = df[['time', 'AtmosphericPressure', 'WindDirection',
                        'WindSpeed', 'Gust', 'WaveHeight', 'WavePeriod',
                        'MeanWaveDirection', 'AirTemperature',
                        'SeaTemperature', 'RelativeHumidity']]

        session_log_data.append(['Data Validation Started <<<<<<<<<<'])
        session_log_data.append([str(pd.Timestamp.now())])

        # Validate for Missing Values
        print("Validating missing values started       <<<<<\n")
        session_log_data.append(['Checking For Missing Values'])
        session_log_data.append([str(pd.Timestamp.now())])

        pd.set_option('future.no_silent_downcasting', True)
        master_df = master_df.replace(to_replace=['nan', 'NaN', ''],
                                      value=np.nan)
        missing_values = master_df.isnull().sum()
        # If there are missing values, write to the session log and error log
        if missing_values.any():
            print("     We found rows with missing values")
            print("     Please check the error log")
            session_log_data.append(
                ['We found missing values in the master data'])
            session_log_data.append([str(pd.Timestamp.now())])
            error_log_data.append(['Missing Values       <<<<<'])
            # append missing values in a column
            for column, count in missing_values.items():
                if count > 0:
                    error_log_data.append(
                        [f"{column}: {count} missing values"])
        else:
            print("     There were no row swith missing data\n")
            session_log_data.append(
                ['We found no missing values in the master data'])

        # Remove rows with missing values
        if missing_values.any():
            session_log_data.append(
                ['Removing data with no values from master data'])
            session_log_data.append([str(pd.Timestamp.now())])
            missing_values_removed_df = master_df.dropna()
            missing_values = missing_values_removed_df.isnull().sum()
            print("     Rows with missing data have been removed\n")

        print("Validating missing values completed     <<<<<\n\n\n")

        # Check for duplicate rows
        print("Validating duplicates started       <<<<<\n")
        duplicates_found = missing_values_removed_df.duplicated(
            keep=False).sum()
        if duplicates_found:
            print(
                f"     There are {duplicates_found} duplicates in the  "
                f"working data set")
            print("     Please check the error log")
            session_log_data.append(
                [f'Duplicates found in data set: {duplicates_found}'])
            session_log_data.append([str(pd.Timestamp.now())])
            error_log_data.append(['Duplicate Rows Found'])
            duplicates_df = missing_values_removed_df[
                missing_values_removed_df.duplicated(keep=False)]

            # Format duplicates_df for column-wise insertion
            duplicates_list_of_lists = duplicates_df.values.tolist()
            # Add header for duplicates (optional)
            duplicates_header = [duplicates_df.columns.tolist()]
            # Combine header and data
            formatted_duplicates = duplicates_header + duplicates_list_of_lists
            # Append to the error log data
            error_log_data.append(['Duplicate Rows Data'])
            error_log_data.extend(formatted_duplicates)
            no_duplicates_df = missing_values_removed_df.drop_duplicates(
                keep='first')
            print("     Duplicates have been removed\n")
        else:
            print("     No duplicates found in the working data set.\n")
            no_duplicates_df = missing_values_removed_df

        print("Validating duplicates completed     <<<<<\n\n\n")

        # Check for outliers
        print("Outlier Validation Started       <<<<<\n")
        session_log_data.append(['Outlier Validation Started'])
        session_log_data.append([str(pd.Timestamp.now())])

        numeric_df = no_duplicates_df.apply(
            lambda col: col.map(lambda x: pd.to_numeric(x, errors='coerce')))

        atmospheric_outliers = check_for_outliers(
            numeric_df[['AtmosphericPressure']])
        wind_outliers = check_for_outliers(numeric_df[['WindSpeed', 'Gust']])
        wave_outliers = check_for_outliers(
            numeric_df[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']])
        temp_outliers = check_for_outliers(
            numeric_df[['AirTemperature', 'SeaTemperature']])

        # Update Outlier Sheets
        if not atmospheric_outliers.empty:
            atmos_outlier_log.update(df_to_list_of_lists(atmospheric_outliers),
                                     'A1')
            print(
                "     Atmospheric Outliers Were Found: Check Atmos Outlier Log")
            print(
                f"    \nThe link to the Atmospheric Outliers log is: \n\
                \n{atmos_outliers_url}\n\n")
        if not wind_outliers.empty:
            wind_outlier_log.update(df_to_list_of_lists(wind_outliers), 'A1')
            print("     Wind Outliers Were Found: Check Atmos Outlier Log")
            print(
                f"    \nThe link to the Wind Outliers log is:\n\n"
                f"{wind_outliers_url}\n\n")
        if not wave_outliers.empty:
            wave_outlier_log.update(df_to_list_of_lists(wave_outliers), 'A1')
            print(
                "     Wave Outliers Were Found:        Check Wave Outlier Log")
            print(
                f"    \nThe link to the Wave Outliers log is:\n\n"
                f"{wave_outliers_url}\n\n")
        if not temp_outliers.empty:
            temp_outlier_log.update(df_to_list_of_lists(temp_outliers), 'A1')
            print(
                "     Temp Outliers Were Found:        Check Temp  Outlier "
                "Log\n")
            print(
                f"    \nThe link to the Temp Outliers log is:\n\n"
                f"{temp_outliers_url}\n\n")

        print("Outlier Validation Completed     <<<<<\n\n\n")

        # Check for date inconsistencies
        print("Date Validation Started       <<<<<\n")
        session_log_data.append(['Date and Time Validation Started'])
        session_log_data.append([str(pd.Timestamp.now())])
        date_time_pattern = r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
        inconsistent_date_format = no_duplicates_df[
            ~no_duplicates_df['time'].str.match(date_time_pattern)]

        if inconsistent_date_format.empty:
            print("     No Incorrect Date Formats Found")
            validated_data_df = no_duplicates_df.copy()
            validated_data_df['time'] = pd.to_datetime(
                validated_data_df['time'], format='%Y-%m-%dT%H:%M:%SZ')
            validated_data_df['time'] = validated_data_df['time'].dt.strftime(
                '%d-%m-%YT%H:%M:%S')
        else:
            print("     Incorrect Date Formats Found\n")
            date_time_log_data.append(['Inconsistent Date and Time Formats'])
            date_time_log_data.append(
                inconsistent_date_format['time'].fillna('').astype(
                    str).values.tolist())
            validated_data_df = no_duplicates_df[
                no_duplicates_df['time'].str.match(date_time_pattern)].copy()
            validated_data_df['time'] = pd.to_datetime(
                validated_data_df['time'], format='%Y-%m-%dT%H:%M:%SZ')
            validated_data_df['time'] = validated_data_df['time'].dt.strftime(
                '%d-%m-%YT%H:%M:%S')

        print("     Incorrect Date Formats Removed\n")
        print("Date Validation Completed     <<<<<\n")
        session_log_data.append(['Data Validation Ended <<<<<<<<<<'])
        session_log_data.append([str(pd.Timestamp.now())])

        # Convert all elements in logs to strings
        session_log_data = [[str(item) for item in sublist] for sublist in
                            session_log_data]
        error_log_data = [[str(item) for item in sublist] for sublist in
                          error_log_data]
        date_time_log_data = [[str(item) for item in sublist] for sublist in
                              date_time_log_data]

        # Write all accumulated data at once
        # Log inconsistent date formats here
        date_time_log.update(date_time_log_data, 'A1')
        print(
            f"    \nDate Inconsistancies found are written here: \n\
            \n{date_time_url}\n\n")


    except Exception as e:
        session_log_data.append([f"An error occurred during validation: {e}"])
        print(f"An error occurred during validation: {e}")

    finally:
        # Convert log lists to strings and update Google Sheets
        session_log.update(df_to_list_of_lists(pd.DataFrame(session_log_data)),
                           'A1')
        error_log.update(df_to_list_of_lists(pd.DataFrame(error_log_data)),
                         'A1')
        date_time_log.update(
            df_to_list_of_lists(pd.DataFrame(date_time_log_data)), 'A1')

    print("\n\n\n >>>>> Master Data Validation Completed <<<<<\n\n\n")
    print("Writing Validated Data To Google Sheets Started      <<<<<\n")
    set_with_dataframe(validated_master_data, validated_data_df)
    print("Writing Validated Data To Google Sheets Completed    <<<<<\n")
    print(
        f"    \nThe validated master data is written here:\n\n"
        f"{validated_master_data_ulr}\n\n")
    print("\n")
    print("#############################################################")
    print("   >>>>>   Phase 1. Completed Data Validation Process   <<<<<")
    print("#############################################################")
    print("\n")

    return validated_data_df


def format_df_date(validated_df):
    """
    Function to change the format of the time column date
    from yyyy-mm-dd to dd-mm-yyyy
    as this is the format the user will input as their
    reqesed date range
    """
    # convert the time columnin df to just date for comparison drop h:m:s
    validated_df['time'] = pd.to_datetime(validated_df['time'],
                                          format='%d-%m-%YT%H:%M:%S',
                                          errors='coerce', dayfirst=True)
    # Extract the date component only, ignoring the time
    validated_df['date_only'] = validated_df['time'].dt.strftime('%d-%m-%Y')


def validate_input_dates(date_str, reference, df_first_date, df_last_date):
    """
    Validates an input date string to ensure it is in the correct format, 
    represents a valid calendar date, and falls within the specified date range.

    This function performs the following tasks:
    1. Converts the input date string `date_str` to a `datetime` object.
    2. Checks that the date is in 'dd-mm-yyyy' format and is a valid date.
    3. Validates the day, month, and year individually, ensuring:
       - The month is between 1 and 12.
       - The day is within the correct range for the specified month 
         (taking into account leap years for February).
       - The date falls within the available date range specified by 
         `df_first_date` and `df_last_date`.
    4. Returns the validated date as a `datetime` object if all checks pass.
    5. Raises a `ValueError` if the date is invalid, out of range, 
       or incorrectly formatted.

    Args:
    - date_str (str): The date string to be validated, expected in 
      'dd-mm-yyyy' format.
    - reference (str): A string indicating whether the date is a "start" or 
      "end" date, used for error messaging.
    - df_first_date (str): The earliest available date in the dataset, 
      in 'dd-mm-yyyy' format.
    - df_last_date (str): The latest available date in the dataset, 
      in 'dd-mm-yyyy' format.

    """
    try:
        # Convert the input to a datetime object
        date = pd.to_datetime(date_str, format='%d-%m-%Y')
        if pd.isna(date):
            raise ValueError(
                "Date is not in the correct format or is invalid.")

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
                    raise ValueError(
                        "February in a leap year has only 29 days.")
            else:
                if day > 28:
                    raise ValueError(
                        "February has only 28 days in a non-leap year.")

        # Ensure the date falls within the available range
        if date < pd.to_datetime(df_first_date) or date > pd.to_datetime(
                df_last_date):
            raise ValueError(
                f"Date {date_str} is outside the available data range.")
        return date
    except ValueError as ve:
        # Re-raise the error with a custom message
        raise ValueError(
            f"Invalid date: {ve}. Please enter the date in  'dd-mm-yyyy' "
            f"format.")


def get_user_dates(validated_df):
    """
    This function:
    - displays the date range available in master data
    - asks for input start date
    - asks for input end date

    Args:
    - validated_df (pd.DataFrame): A DataFrame containing a 'time' 
      column with date and time data.
    """
    # Declare variable for later use
    error_log_data = []

    # Convert the time column in validated_df to date time
    validated_df['time'] = pd.to_datetime(validated_df['time'],
                                          format='%d-%m-%YT%H:%M:%S')

    # Get the first and last dates available in the data frame
    df_first_date = validated_df['time'].min().strftime('%d-%m-%Y')
    df_last_date = validated_df['time'].max().strftime('%d-%m-%Y')
    print(f"\n\nAvailable data range: {df_first_date} to {df_last_date}\n\n")
    print("You will be asked to enter a start date and and end date")
    print(">>> You can type 'quit' at any time to exit. <<<\n")

    # Prompt user for start date, dont go to end date until date is acceptable and formatted
    while True:
        user_input_start_date = input(
            f"Please enter the start date\n (in the format 'dd-mm-yyyy' "
            f"\nbetween {df_first_date} and {df_last_date}): ")
        # Check to see if the user wants to quit
        if user_input_start_date.lower() == 'quit':
            print("Exiting program as requested.")
            exit()

        try:
            user_input_start_date = validate_input_dates(user_input_start_date,
                                                         "start",
                                                         df_first_date,
                                                         df_last_date)
            break  # Exit loop if the date is valid
        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(
                ["Start Date Error", str(pd.Timestamp.now())])
            error_log_data.append(["Start Date Error", str(e)])
            print("Start Date Error:\n")
            print(f"You Entered: {user_input_start_date}  <<<<<\n")
            print(
                f"\n A detailed description of the error\n has been  "
                f"appended to the error log.")
            print("\nPlease enter the date in 'dd-mm-yyyy' format.\n\n")

    # Prompt user for end date
    while True:
        user_input_end_date = input(
            f"\n\nPlease enter the end date\n (in the format 'dd-mm-yyyy'\n  "
            f"within {df_first_date} and {df_last_date}): ")
        # Check to see if the user wants to quit
        if user_input_end_date.lower() == 'quit':
            print("Exiting program as requested.")
            exit()

        try:
            user_input_end_date = validate_input_dates(user_input_end_date,
                                                       "end", 
                                                       df_first_date,
                                                       df_last_date)
            if user_input_end_date < user_input_start_date:
                raise ValueError(
                    "The end date cannot be earlier than the start date.")
            break  # Exit loop if the date is valid
        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(["End Date Error", str(pd.Timestamp.now())])
            error_log_data.append(["End Date Error", str(e)])
            print(f"End Date Error:\n")
            print(f"You Entered: {user_input_start_date}  <<<<<\n")
            print(
                f"\n A detailed description of the error\n has been appended "
                f"to the error log.")
            print("\nPlease enter the date in 'dd-mm-yyyy' format.\n\n")

    # Convert validated start and end dates to string format 
    # for further processing
    user_input_start_date_str = user_input_start_date.strftime('%d-%m-%Y')
    user_input_end_date_str = user_input_end_date.strftime('%d-%m-%Y')

    # Write any errors to log
    log_errors_to_sheet(error_log_data)

    return user_input_start_date_str, user_input_end_date_str


def filter_data_by_date(validated_df, start_date, end_date):
    """
    Filter the DataFrame to include only rows within the specified date range.

    This function filters the input DataFrame based on a date range defined by 
    `start_date` and `end_date`.
    It assumes the DataFrame contains a 'date_only' column with dates formatted
    as 'day-month-year'.

    Args:
    - validated_df (pd.DataFrame): The input DataFrame containing a 'date_only'
    - start_date (pd.Timestamp or str): The start date of the range to filter.
    - end_date (pd.Timestamp or str): The end date of the range to filter.
    """
    return validated_df[
        (pd.to_datetime(validated_df['date_only'],
                        format='%d-%m-%Y') >= start_date) &
        (pd.to_datetime(validated_df['date_only'],
                        format='%d-%m-%Y') <= end_date)
        ]


def format_df_data_for_display(date_filtered_df):
    """
    Format the DataFrame for display by converting the 'time' column to a 
    specific datetime format.

    This function takes a DataFrame with a 'time' column, creates a copy of it,
    and formats the 'time'
    column into a string representation with the format 'day-month-year 
    hour:minute:second'. This ensures
    that the datetime values are presented in a user-friendly format suitable 
    for display purposes.

    Args:
    - date_filtered_df (pd.DataFrame): The input DataFrame containing at least 
    a 'time' column with datetime values.
    """
    # Create a copy of date_filtered_df to avoid modifying the original 
    # DataFrame
    working_data_df = date_filtered_df.copy()

    # Convert the 'time' column to the desired format in the copied DataFrame
    working_data_df['time'] = pd.to_datetime(
        working_data_df['time']).dt.strftime('%d-%m-%Y %H:%M:%S')

    return working_data_df


def get_data_selection():
    """
    Prompts the user to select data columns to display from a predefined 
    set of options.

    The function presents a menu of options for selecting different 
    types of data to display. The user can choose from:
    1. All Data
    2. Atmospheric Pressure
    3. Wind Speed and Gust
    4. Wave Height, Wave Period, and Mean Wave Direction
    5. Air Temperature and Sea Temperature
    6. Exit Output Options

    The function will continuously prompt the user until a valid selection 
    is made or the user opts to exit.
    If the user inputs an invalid selection, an error message is displayed, 
    and the error details are logged.
    """
    # Initialise selection storage
    selected_columns = []
    error_log_data = []
    while True:
        # Output Selection Options
        print("\nSelect the data you want to display:")
        print("1: All Data")
        print("2: Atmospheric Pressure")
        print("3: Wind Speed and Gust")
        print("4: Wave Height, Wave Period, and Mean Wave Direction")
        print("5: Air Temperature and Sea Temperature")
        print("6: Exit Ouput Options\n")

        # take users selected number
        selection = input("Enter the number corresponding to your selection: ")
        try:
            # convert selection to integer for use later
            selection = int(selection)

            # Create relevant data subset dataframes for processing
            if selection == 1:
                selected_columns = ['time', 'AtmosphericPressure',
                                    'WindDirection',
                                    'WindSpeed', 'Gust', 'WaveHeight',
                                    'WavePeriod',
                                    'MeanWaveDirection', 'AirTemperature',
                                    'SeaTemperature', 'RelativeHumidity']
            elif selection == 2:
                selected_columns = ['time', 'AtmosphericPressure']
            elif selection == 3:
                selected_columns = ['time', 'WindSpeed', 'Gust']
            elif selection == 4:
                selected_columns = ['time', 'WaveHeight', 'WavePeriod',
                                    'MeanWaveDirection']
            elif selection == 5:
                selected_columns = ['time', 'AirTemperature', 'SeaTemperature']
            elif selection == 6:
                return []  # Exit the function, returning an empty list.
            else:
                raise ValueError(
                    "Selection out of range. Please select a number between "
                    "1 and 6.")

            # If a valid selection is made, break out of the loop and 
            # return the selected columns
            break

        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(
                ["Output Selection Error", str(pd.Timestamp.now())])
            error_log_data.append(["You Input ", selection])
            error_log_data.append(["Error Description", str(e)])
            print(f"Output Selection Error:\n")
            print(f"You Entered: {selection}    <<<<<\n")
            print(
                f"Invalid data selection input.\n A detailed description "
                f"of the error\n has been appended to the error log.")
            print("\nPlease enter 1, 2, 3, 4, 5, 6 to exit\n")

        # Write any errors to log
        log_errors_to_sheet(error_log_data)

    return selected_columns


def determine_output_options(num_rows):
    """
    Determine the allowed output options based on the number of rows 
    in a dataset.

    Args:
    - num_rows (int): The number of rows in the dataset.

    Returns:
    tuple: A tuple containing three boolean values:
        - allow_screen (bool): Whether to allow displaying the 
          data on the screen.
        - allow_graph (bool): Whether to allow generating graphical 
          representations (e.g., charts).
        - allow_sheet (bool): Whether to allow exporting the data to 
          a spreadsheet.
    """
    # Check to see how many rows there are in the user output dataframe
    if num_rows <= 30:
        # If there are 30 or fewer rows, allow all actions
        allow_screen = True
        allow_graph = True
        allow_sheet = True
    else:
        # If more than 20 rows, limit options
        allow_screen = False
        allow_graph = True
        allow_sheet = True
        print(
            f"\nNote: Displaying the first 30 rows. There are {num_rows - 30} "
            f"more rows not displayed.")

    return allow_screen, allow_graph, allow_sheet


def get_output_selection(user_output_df, selected_columns, allow_screen,
                         allow_graph, allow_sheet, num_rows):
    """
    Prompts the user to select an output option for a DataFrame and performs 
    the corresponding action.

    This function provides an interactive loop that allows the user to choose 
    how they want to handle the output
    of the provided DataFrame (`user_output_df`). The options include 
    displaying data on the screen, plotting a graph
    in the browser, writing data to a Google Sheet, or exiting the loop.

    Args:
    - user_output_df (pd.DataFrame): The DataFrame to be processed and 
      outputted based on user selection.
    - selected_columns (list of str): The list of column names available in 
      `user_output_df` for graphing or other operations.
    - allow_screen (bool): Flag indicating whether the option to display data 
      on the screen is available.
    - allow_graph (bool): Flag indicating whether the option to generate a 
      graph is available.
    - allow_sheet (bool): Flag indicating whether the option to write data to 
      a Google Sheet is available.
    - num_rows (int): Number of rows in `user_output_df`. Used to determine 
      how many rows to display on the screen.
    """

    # Inner loop allowing user select different output options
    while True:
        # Get the action from user with validation
        output_selection = get_valid_data_output_selection(allow_screen,
                                                           allow_graph,
                                                           allow_sheet)
        try:
            # If user selects 1 - output to screen
            if output_selection == 1:
                # Print the first 20 rows if more than 20
                print("\nSelected Data:")
                if num_rows > 30:
                    print(user_output_df.head(30))
                else:
                    print(user_output_df)
            # If user selects 2 - Output goes to graph in browser
            elif output_selection == 2:
                # Option 2: Output to graph
                # Plotting weather data over time
                x_col = 'time'
                # Refactored to select columns except the time column 
                # for y axis
                y_cols = [col for col in selected_columns if col != x_col]
                title = 'Weather Data Over Time'
                user_requested_graph(user_output_df, x_col, y_cols, title)
            # If user selects 3 - output to google sheet
            elif output_selection == 3:
                # Option 3 Write Data To Google Sheet
                set_with_dataframe(user_data_output, user_output_df)
                print("\nData Written To Google Sheet")
                print(
                    f"    \nData Output Can Be Found Here:"
                    f"\n\n{user_data_output_url}\n\n")
            # If user select 4 - loop ends
            elif output_selection == 4:
                # Option 4 Exit the loop
                print("\nExited Output Options ......\n")
                break

        except ValueError as e:
            print("Error in output selection")


def data_initialisation_and_validation():
    """
    This function:
    - Clears all the sheets
    - loads the marine data for validation
    - Validates the data
    - Stores the validated data for use in the session
    Its only run once per session
    """
    print("\n")
    print("#############################################################")
    print(">>>>>   Phase 1. Starting Data Validation Process       <<<<<")
    print("#############################################################")
    print("\n")
    # Load the marine data for validation
    master_data, session_log_data, error_log_data =  (
        load_marine_data_input_sheet())
    # Create a validated data frame for use in the app
    validated_df = validate_master_data(master_data, session_log_data,
                                        error_log_data)

    return validated_df


def convert_dataframe(df, x_col, y_cols):
    """
    Convert columns in DataFrame to appropriate data types.

    Args:
    - df : The DataFrame containing the data to be converted.
    - x_col : The name of the column to be converted to datetime format.
    - y_cols : A list of column names to be converted to numeric format.
    """
    # Convert the x_col to datetime if it represents dates
    df.loc[:, x_col] = pd.to_datetime(df[x_col], errors='coerce')

    # Convert y_cols to numeric
    for col in y_cols:
        df.loc[:, col] = pd.to_numeric(df[col], errors='coerce')

    return df


def write_data_to_sheet(service, spreadsheet_id, sheet_name, values):
    """
    Write data to the specified Google Sheet.

    Args:
    - service: Authorized Google Sheets API service instance.
    - spreadsheet_id: The ID of the Google Spreadsheet.
    - sheet_name: The name of the sheet within the spreadsheet.
    - values: A list of lists, where each inner list represents a 
      row of data to be written.
    """
    try:
        body = {'values': values}
        range_ = f'{sheet_name}!A1'
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_,
            valueInputOption='RAW',
            body=body
        ).execute()
    except HttpError as e:
        print(f"HTTP error occurred while writing data: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred while writing data: {e}")


def delete_existing_charts(service, spreadsheet_id):
    """
    Check if there are existing charts in the specified Google Sheet and 
    delete them if found.

    Args:
    - service: Google Sheets API service instance.
    - spreadsheet_id: ID of the spreadsheet.
    """
    try:
        # Retrieve the spreadsheet information
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id).execute()

        # Extract sheet information
        sheets = spreadsheet.get('sheets', [])
        charts_to_delete = []

        # Iterate over sheets to find charts
        for sheet in sheets:
            sheet_id = sheet.get('properties', {}).get('sheetId')
            sheet_name = sheet.get('properties', {}).get('title')
            if 'charts' in sheet:
                for chart in sheet.get('charts', []):
                    chart_id = chart.get('chartId')
                    charts_to_delete.append(chart_id)

        # Delete charts if any are found
        if charts_to_delete:
            requests = [{'deleteEmbeddedObject': {'objectId': chart_id}} for
                        chart_id in charts_to_delete]
            batch_update_request = {'requests': requests}
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=batch_update_request
            ).execute()

            print(
                f"Deleted {len(charts_to_delete)} charts from "
                f"the Google Sheet.")
        else:
            print("No charts found to delete.")

    except HttpError as e:
        print(f"HTTP error occurred: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def add_chart_to_sheet(service, spreadsheet_id, sheet_id, x_col, y_cols,
                       title):
    """
    Add a chart to the Google Sheet with the legend positioned on the right,
    and configure the chart to use the first row as headers.

    Args:
    - service: Google Sheets API service instance.
    - spreadsheet_id: ID of the spreadsheet.
    - sheet_id: ID of the sheet within the spreadsheet.
    - x_col: The header name or index of the column to be used for the x-axis.
    - y_cols: List of header names or indices of the columns to be used for 
      the y-axis (multiple series).
    - title: Title of the chart.
    """
    delete_existing_charts(service, spreadsheet_id)
    try:
        # Determine the data range for the chart
        # Assuming that the number of rows matches the length of y_cols
        end_row_index = len(
            y_cols) + 1  

        # Create the requests to add the chart
        # Define your chart specifications
        requests = [
            {
                'addChart': {
                    'chart': {
                        'spec': {
                            'title': title,
                            'basicChart': {
                                'chartType': 'LINE',
                                'legendPosition': 'RIGHT_LEGEND',
                                # Position the legend on the right
                                'headerCount': 1,
                                # This sets the first row as headers
                                'axis': [
                                    {'position': 'BOTTOM_AXIS',
                                     'title': x_col},  # X-axis title
                                    {'position': 'LEFT_AXIS',
                                     'title': 'Values'}  # Y-axis title
                                ],
                                'domains': [
                                    {
                                        'domain': {
                                            'sourceRange': {
                                                'sources': [
                                                    {
                                                        'sheetId': sheet_id,
                                                        'startRowIndex': 0,
                                                        # Data starts in 
                                                        # the header
                                                        'endRowIndex': 
                                                            end_row_index,
                                                        'startColumnIndex': 0,
                                                        'endColumnIndex': 1
                                                    }
                                                ]
                                            }
                                        }
                                    }
                                ],
                                'series': [
                                    {
                                        'series': {
                                            'sourceRange': {
                                                'sources': [
                                                    {
                                                        'sheetId': sheet_id,
                                                        'startRowIndex': 0,
                                                        # Data starts in 
                                                        # the header
                                                        'endRowIndex': 
                                                            end_row_index,
                                                        'startColumnIndex': 
                                                            col_index,
                                                        'endColumnIndex': 
                                                            col_index + 1
                                                    }
                                                ]
                                            }
                                        }
                                    } for col_index in
                                    range(1, len(y_cols) + 1)
                                ]
                            }
                        },
                        'position': {
                            'overlayPosition': {
                                'anchorCell': {
                                    'sheetId': sheet_id,
                                    'rowIndex': 0,
                                    'columnIndex': len(y_cols) + 2
                                    # Adjust to place the chart appropriately
                                },
                                'offsetXPixels': 0,
                                'offsetYPixels': 0
                            }
                        }
                    }
                }
            }
        ]

        # Execute the batch update request
        batch_update_request = {'requests': requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=batch_update_request
        ).execute()

        print("Chart has been created in the Google Sheet")
        print(
            f"    \nYou Can View Your Chart Here:"
            f"\n\n{graphical_output_data_url}\n\n")
    except HttpError as e:
        print(f"HTTP error occurred while creating the chart: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred while creating the chart: {e}")


def user_requested_graph(df, x_col, y_cols, title):
    """
    Writes data from a Pandas DataFrame to a specified Google Sheet and 
    creates a chart.

    This function takes a Pandas DataFrame, extracts specified columns, 
    and writes the data
    to a Google Sheet. It then creates a chart in the Google Sheet using the 
    specified columns.

    Args:
    - df (pd.DataFrame): The DataFrame containing the data to be written 
    to the Google Sheet.
    - x_col (str): The name of the column in `df` to be used as the x-axis 
    in the chart.
    - y_cols (list of str): A list of column names in `df` to be used as the 
    y-axis in the chart.
    - title (str): The title of the chart to be created in the Google Sheet
    """
    try:
        # Define the Google Sheet ID and the range where data will be written
        spreadsheet_id = '1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U'
        sheet_name = 'graphical_output_data'
        sheet_id = 2079660690  # The specific sheet ID of the output sheet
        service = build('sheets', 'v4', credentials=SCOPED_CREDS)

        # Convert the DataFrame to correct data types
        df = convert_dataframe(df, x_col, y_cols)

        # Prepare data to be written to the sheet
        values = [df.columns.tolist()]  # Header row
        for index, row in df.iterrows():
            values.append(
                [row[x_col].strftime('%Y-%m-%d')] + [row[col] for col in
                                                     y_cols])

        # Write data to the sheet
        write_data_to_sheet(service, spreadsheet_id, sheet_name, values)

        # Add a chart to the sheet
        add_chart_to_sheet(service, spreadsheet_id, sheet_id, x_col, y_cols,
                           title)

    except HttpError as e:
        print(f"HTTP error occurred: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def get_valid_data_output_selection(allow_screen, allow_graph, allow_sheet):
    """
    Function to prompt the user for output selection and ensure that only 
    valid numbers are allowed.
    Its been modified for conditional logic to check if the output is
    too large for the screen. Conidered > 30 rows of data

    - allow_screen controls if user can print to screen
    - allow_graph controls if user can send datat to graph
    - allow_sheet controls if usre can write to sheet
    """
    error_log_data = []
    while True:
        # Provide user options for output
        print("\nWhat would you like to do with the selected data?")
        if allow_screen:
            print("1: Print to Screen")
        if allow_graph:
            print("2: Create Graph")
        if allow_sheet:
            print("3: Write to Google Sheet")
        print("4: Exit")

        user_input = input(
            "\nEnter the number corresponding to your desired output: ")

        try:
            # Attempt to convert input to an integer
            output_selection = int(user_input)
            # Check if the number is within the valid range and allowed
            if output_selection in [1, 2, 3, 4] and \
                    ((output_selection == 1 and allow_screen) or
                     (output_selection == 2 and allow_graph) or
                     (output_selection == 3 and allow_sheet) or
                     output_selection == 4):
                return output_selection  # Return the valid selection

            else:
                print("\n\nInvalid selection. Please enter a valid option.\n")

        except ValueError as e:
            # Log the error
            error_log_data.append(
                ["Output Selection Error", str(pd.Timestamp.now())])
            error_log_data.append(["Output Selection Error", str(e)])
            print(f"Output Selection Error:\n")
            print(f"You Entered: {user_input}  <<<<<\n")
            print(
                f"\nA detailed description of the error\nhas been appended "
                f"to the error log.")
            print("Invalid selection. Please enter a number between 1 and 4.")

        # Write any errors to log
        log_errors_to_sheet(error_log_data)

    return output_selection


def log_errors_to_sheet(error_log_data):
    """
    Write errors to the Google Sheet, starting at cell A25 and appending 
    subsequent errors.

    Args:
    - error_log_data : A list where each item is a tuple containing error 
      information.
    - error_log : The Google Sheets worksheet object where errors will 
      be logged.

    """
    # Determine where to start appending errors
    start_row = 4
    existing_values = error_log.get_all_values()
    if existing_values:
        last_row = len(existing_values) + start_row - 1
    else:
        last_row = start_row - 1

    # Append new errors to the sheet
    for i, error in enumerate(error_log_data):
        row = last_row + i + 1
        # Ensure the tuple has at least two elements
        error_id = error[0] if len(error) > 0 else "No ID provided"
        error_message = error[1] if len(
            error) > 1 else "No error message provided"

        # Update columns A and B
        error_log.update_cell(row, 1, error_id)  # Column A
        error_log.update_cell(row, 2, error_message)  # Column B

    # Clear the error_log_data list after logging
    error_log_data = []


def get_continue_yn():
    """
    Introduces the application to the user and prompts them to decide 
    whether they want to continue.

    The function displays an introduction message about the application, 
    which consists of two parts:
    - Data Validation
    - Data Interrogation

    It then enters a loop asking the user if they want to continue. 
    The user should input 'y' for yes or 'n' for no.
    If the user inputs 'y', the function returns 'y' and continues. 
    If the user inputs 'n', the application exits.
    If the user inputs an invalid response, an error message is displayed, 
    logged, and the prompt is repeated.

    The function logs any errors related to invalid user input and provides 
    feedback to the user accordingly.

    Returns:
        str: The user's response ('y' to continue or 'n' to exit).

    """
    error_log_data = []
    # Display app introduction
    print("\n\nWelcome to the Weather Data Analysis Application.\n")
    print(
        "This application has two parts:\n - Data Validation\n - "
        "Data Interrogation.\n")

    # Start the query loop
    # Start the query loop
    while True:
        proceed = input("Do you want to continue? (y/n): ").strip().lower()

        try:
            # Check user input
            if proceed == 'y':
                return proceed
            elif proceed == 'n':
                print("Exiting Application... Good Bye")
                exit()
            else:
                raise ValueError("Selection must be (y/n)")

        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(
                ["User Input Error", str(pd.Timestamp.now())])
            error_log_data.append(["You Entered", str(proceed)])
            error_log_data.append(["User Input Error", str(e)])
            print(f"User Input Error:\n")
            print(f" Error: {e}\n")
            print(
                f"Invalid input.>>> {proceed} <<<\n\nA detailed description "
                f"of the error\n has been appended to the error log.")
            print("\nPlease enter (n) to exit\n")

        # Write any errors to log
        log_errors_to_sheet(error_log_data)


def main():
    """
    Main function that controls all the functionality of the application

    This function performs the following tasks:
    1. Prompts the user to continue or exit the application.
    2. Initializes and validates the master data.
    3. Checks if data validation was successful and proceeds if valid.
    4. Formats the validated data frame's date columns.
    5. Allows the user to specify a date range for filtering the data.
    6. Filters the data frame based on the user-specified date range.
    7. Formats the filtered data frame for display purposes.
    8. Provides options for users to select specific data columns for output.
    9. Determines and manages output options based on the number of rows 
       in the data.
    10. Generates output based on user preferences and choices.
    11. Logs and writes error data to an error log sheet if applicable.
    """
    error_log_data = []
    # Initialise Sheets On Load
    initialise_google_sheets()

    while True:
        # Introduce the app and ask user if they want to continue
        decision = get_continue_yn()

        # Initialise the sheets and validate the data
        validated_df = data_initialisation_and_validation()

        # Check to see if the data has been validated
        if validated_df is None:
            print(
                "Error: Data has not been validated\n The application cannot "
                "continue..... \nBye...")
            exit()
        else:
            print("\n")
            print("#########################################################")
            print("    >>>>> Phase 2. Data Interrogation Starting      <<<<<")
            print("#########################################################")
            print("\n")
            print("Please Check The Error Log Files For Any Errors Found")
            print(
                f"    \nSession Errors Can Be Found Here:"
                f"\n\n{gael_force_error_log_url}\n\n")
        # Convert the data frame date format to dd-mm-yyyy
        format_df_date(validated_df)

        # Outer Loop - get dates from user for specified range
        while True:
            # Get dates from user to interrogate validated data frame
            user_input_start_date_str, user_input_end_date_str = (
                get_user_dates(
                validated_df))

            # Check if user selected "quit"
            if (user_input_start_date_str is None or user_input_end_date_str 
                    is None):
                print("\n\nDate input was canceled. Exiting program.\n\n")
                exit()  # Exit app

            # Convert user input dates to the format used in validated_df
            user_input_start_date = user_input_start_date_str
            user_input_end_date = user_input_end_date_str
            print(f"Date Range - Beginning: {user_input_start_date}")
            print(f"Date Range - End: {user_input_end_date}")

            # Filter the dataframe based on the date range
            date_filtered_df = filter_data_by_date(validated_df,
                                                   user_input_start_date,
                                                   user_input_end_date)

            # Format the dataframe for display, converting date format 
            # to dd-mm-yyyy
            working_data_df = format_df_data_for_display(date_filtered_df)
            # Middle loop - getting specific data set for user output
            while True:
                # Get users selection for data output columns
                selected_columns = get_data_selection()
                if not selected_columns:
                    break
                user_output_df = working_data_df[selected_columns]
                num_rows = len(user_output_df)
                print(f"\nThere are {num_rows} rows of data.     <<<<<\n")

                # Determine output options based on rows
                allow_screen, allow_graph, allow_sheet = (
                    determine_output_options(
                    num_rows))

                # Create output based on user selection
                get_output_selection(user_output_df, selected_columns,
                                     allow_screen, allow_graph, allow_sheet,
                                     num_rows)

    # Write error log to error log sheet if there are errors to be written
    if error_log_data:
        error_log_data = [[str(item) for item in sublist] for sublist in
                          error_log_data]
        error_log.update(df_to_list_of_lists(pd.DataFrame(error_log_data)),
                         'A25')


main()
