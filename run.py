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


# Define the function to check Google Sheet access
def check_google_sheet_access(credentials_path, sheet_name):
    """
    Verifies access to a Google Sheet using the specified service
    account credentials.

    Args:
    credentials_path (str): The file path to the JSON file containing
    the service account credentials.
    sheet_name (str): The name of the Google Sheet to be accessed.

    Returns:
    gspread.models.Worksheet: The first worksheet object if the Google Sheet
    is successfully accessed.
    None: If access to the Google Sheet fails, returns None and prints an
    error message.
    """
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]

    try:
        # Load credentials
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            credentials_path, scope
            )
        client = gspread.authorize(creds)

        # Access the Google Sheet
        sheet = client.open(sheet_name).sheet1  # Access the first sheet
        print("Successfully accessed the Google Sheet!")
        return sheet

    except FileNotFoundError as e:
        print(
            f"Error: The specified credentials file was "
            f"not found.\nDetails: {e}"
            f"\n\nWithout access to the master data the app will not run"
            f"\nApplication Ending......... Bye\n"
            )
        exit()
    except PermissionError as e:
        print(
            f"Error: Permission denied while accessing the credentials "
            f"file.\nDetails: {e}"
            f"\n\nWithout access to the master data the app will not run"
            f"\nApplication Ending......... Bye\n"
            )
        exit()
    except gspread.exceptions.SpreadsheetNotFound as e:
        print(
            f"Error: The specified Google Sheet was not found.\nDetails: {e}"
            f"\n\nWithout access to the master data the app will not run"
            f"\nApplication Ending......... Bye\n"
            )
        exit()
    except gspread.exceptions.APIError as e:
        print(
            f"Error: There was an API error when trying to access the "
            f"Google Sheet.\nDetails: {e}"
            f"\n\nWithout access to the master data the app will not run"
            f"\nApplication Ending......... Bye\n"
            )
        exit()
    except IOError as e:
        print(f"Error: An I/O error occurred.\nDetails: {e}")
        exit()
    except Exception as e:
        print(f"An unexpected error occurred.\nDetails: {e}")
        exit()

    return None


# Define the function to initialize Google Sheets
def initialise_google_sheets():
    """
    Initializes the Google Sheets required for the application by
    clearing existing data.

    This function accesses the specified Google Sheets and clears data from
    the relevant worksheets to prepare them for new input. It sets up
    authorization, defines worksheet variables, and provides URL links
    to each worksheet tab.

    Returns:
    tuple: A tuple containing the initialized Google Sheets variables,
    worksheet objects, and URLs for the relevant worksheet tabs.
    If initialization fails, returns None and prints an error message.
    """
    print("\n#######################################")
    print("Initialising Necessary Files To Run App")
    print("#######################################\n")

    sheet = check_google_sheet_access('creds.json',
                                      'marine_data_m2')

    if sheet is not None:
        print("Starting Google Sheet Initialisation\n")

        # After successful access, declare all the variables
        SCOPE = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
            ]

        # Set up authorization for access to Google Sheets
        CREDS = Credentials.from_service_account_file('creds.json')
        SCOPED_CREDS = CREDS.with_scopes(SCOPE)
        GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
        SHEET = GSPREAD_CLIENT.open('marine_data_m2')

        # Define the variables used to access each of the sheets in
        # the Google Sheet
        unvalidated_master_data = SHEET.worksheet(
            'marine_data_master_data_2020_2024'
            )
        validated_master_data = SHEET.worksheet('validated_master_data')
        user_data_output = SHEET.worksheet('user_data_output')
        session_log = SHEET.worksheet('session_log')
        error_log = SHEET.worksheet('gael_force_error_log')
        date_time_error_log = SHEET.worksheet('date_time_error_log')
        graphical_output_sheet = SHEET.worksheet('graphical_output_data')

        # Define the sheets to be used for outlier output
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

        validated_master_data_url = (
            "Tab: validated_master_data - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing178716219"
        )

        user_data_output_url = (
            "Tab: user_data_output - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing911199205"
        )

        session_log_url = (
            "Tab: session_log - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing410413428"
        )

        gael_force_error_log_url = (
            "Tab: gael_force_error_log - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing1094726921"
        )

        atmos_outliers_url = (
            "Tab: atmos_outliers - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing410485236"
        )

        wind_outliers_url = (
            "Tab: wind_outliers - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing594709402"
        )

        wave_outliers_url = (
            "Tab: wave_outliers - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing691663590"
        )

        temp_outliers_url = (
            "Tab: temp_outliers - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing697757766"
        )

        date_time_url = (
            "Tab: date_time - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing2036646154"
        )

        graphical_output_data_url = (
            "Tab: graphical_output_data - URL: "
            "https://docs.google.com/spreadsheets/d/"
            "1cjDvLdeYgYip8yfg4w531LKcoRlo8t8gb8esrI30H6U/"
            "edit?usp=sharing2079660690"
        )

        # Clear the contents of the sheets
        sheets = [
            validated_master_data, user_data_output, session_log, error_log,
            atmos_outlier_log, wind_outlier_log, wave_outlier_log,
            temp_outlier_log, date_time_error_log, graphical_output_sheet
            ]
        for sheet in sheets:
            sheet.clear()

        print("Finished Google Sheet Initialisation\n")

        # Return the variables
        return (
                SHEET, SCOPED_CREDS, unvalidated_master_data,
                validated_master_data,
                user_data_output, session_log, error_log,
                date_time_error_log, graphical_output_sheet,
                atmos_outlier_log, wind_outlier_log, wave_outlier_log,
                temp_outlier_log, google_worksheet, marine_master_data_url,
                validated_master_data_url, user_data_output_url,
                session_log_url, gael_force_error_log_url, atmos_outliers_url,
                wind_outliers_url, wave_outliers_url, temp_outliers_url,
                date_time_url, graphical_output_data_url)

    else:
        print("Failed to initialize Google Sheets due to an error.")


def load_marine_data_input_sheet(
        session_log_url, gael_force_error_log_url,
        session_log, error_log, unvalidated_master_data
        ):
    """
    Initializes the session and error logs, then loads and returns
    the master data.

    This function logs the start of a session and initializes error tracking by
    creating entries in the session log and error log. It then loads the master
    data from the specified Google Sheet into a dataframe.

    Args:
    session_log_url (str): The URL link to the session log Google Sheet.
    gael_force_error_log_url (str): The URL link to the error log Google Sheet.
    session_log (gspread.models.Worksheet): The Google Sheet worksheet for the
    session log.
    error_log (gspread.models.Worksheet): The Google Sheet worksheet
    for the error log.
    unvalidated_master_data (gspread.models.Worksheet):
    The worksheet containing the unvalidated master data.

    Returns:
    tuple: A tuple containing the loaded master data as a dataframe,
    the session log data, and the error log data.
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

    The function calculates the Z-Score for each value in the DataFrame
    and identifies
    outliers as any value with an absolute Z-Score greater than 3.
    These outliers are
    typically considered to be statistically significant deviations
    from the mean.

    Args:
    - df (pandas.DataFrame): The input DataFrame containing numerical
    data to check for outliers.
    """
    z_scores = zscore(df)  # calculate a z score for all values in dataframe
    abs_z_scores = abs(z_scores)  # take the absolute value
    outliers = (abs_z_scores > 3).any(
        axis=1
        )  # creates a bolean based on whether the abs number is  > 3
    return df[outliers]  # return the dataframe of the outliers identified


def handle_log_update(
        update_function, worksheet, data,
        log_name="", error_log_data=None
        ):
    """
    Handles I/O operations with error checking and logging,
    automatically appending data
    to the end of the sheet.

    Args:
    - update_function (function): Function to perform the I/O operation,
    typically the update method of the worksheet.
    - worksheet (object): Google Sheets worksheet object for log operations.
    - data (list): Data to be written or updated.
    - log_name (str): Name of the log for error messages.
    - error_log_data (list, optional): List to store error messages.
    Defaults to None.
    """
    # Initialize error_log_data if it is None
    if error_log_data is None:
        error_log_data = []

    try:
        # Determine the last row in the log to append data
        last_row = get_last_filled_row(worksheet)

        # Determine the starting cell for new data
        start_cell = f"A{last_row + 1}"

        # Append data to the sheet using the update function
        update_function(data, start_cell)
        print(f"Successfully updated {log_name} at {start_cell}.")
    except Exception as e:
        print(f"Error updating {log_name}: {e}")
        error_log_data.append([f"Error updating {log_name}: {e}"])


def get_last_filled_row(worksheet):
    """
    Retrieves the last filled row in a Google Sheet.

    Args:
    - worksheet (object): Google Sheets worksheet object.

    Returns:
    - int: The index of the last filled row.
    """
    sheet_data = worksheet.get_all_values()  # Get all values from the sheet
    return len(sheet_data)  # The last filled row index is the number of rows


def validate_missing_values(master_df, session_log_data, error_log_data):
    """
    Validates the master data for missing values and logs the process.

    This function checks the master data dataframe for missing values,
    logs the validation process in the session log, and records any errors
    found in the
    error log. If missing values are detected, they are removed from
    the dataframe.

    Args:
    master_df (pandas.DataFrame): The dataframe containing the master
    data to be validated.
    session_log_data (list): The list used to log session activity.
    error_log_data (list): The list used to log errors encountered
    during validation.

    Returns:
    pandas.DataFrame: The cleaned dataframe with rows containing
    missing values removed.
    """
    # Validate for Missing Values
    print("Validating missing values started       <<<<<\n")
    session_log_data.append(['Checking For Missing Values'])
    session_log_data.append([str(pd.Timestamp.now())])

    pd.set_option('future.no_silent_downcasting', True)
    master_df = master_df.replace(
        to_replace=['nan', 'NaN', ''],
        value=np.nan
        )
    missing_values = master_df.isnull().sum()
    # If there are missing values, write to the session log and error log
    if missing_values.any():
        print("     We found rows with missing values")
        print("     Please check the error log")
        session_log_data.append(
            ['We found missing values in the master data']
            )
        session_log_data.append([str(pd.Timestamp.now())])
        error_log_data.append(['Missing Values       <<<<<'])
        # append missing values in a column
        for column, count in missing_values.items():
            if count > 0:
                error_log_data.append(
                    [f"{column}: {count} missing values"]
                    )
    else:
        print("     There were no row swith missing data\n")
        session_log_data.append(
            ['We found no missing values in the master data']
            )

    # Remove rows with missing values
    if missing_values.any():
        session_log_data.append(
            ['Removing data with no values from master data']
            )
        session_log_data.append([str(pd.Timestamp.now())])
        missing_values_removed_df = master_df.dropna()
        missing_values = missing_values_removed_df.isnull().sum()
        print("     Rows with missing data have been removed\n")

    print("Validating missing values completed     <<<<<\n\n\n")

    return missing_values_removed_df


def validate_duplicates(
        missing_values_removed_df, session_log_data,
        error_log_data
        ):
    """
    Validates the dataframe for duplicate rows and logs the process.

    This function checks the cleaned dataframe for duplicate rows, logs the
    validation process in the session log, and records any duplicate entries
    found in the error log. If duplicates are detected, they are removed from
    the dataframe.

    Args:
    missing_values_removed_df (pandas.DataFrame): The dataframe that has been
    cleaned of missing values.
    session_log_data (list): The list used to log session activity.
    error_log_data (list): The list used to log errors encountered
    during validation.

    Returns:
    pandas.DataFrame: The dataframe with duplicate rows removed.
    """
    # Check for duplicate rows
    print("Validating duplicates started       <<<<<\n")
    duplicates_found = missing_values_removed_df.duplicated(
        keep=False
        ).sum()

    if duplicates_found:
        print(
            f"     There are {duplicates_found} duplicates in the  "
            f"working data set"
            )
        print("     Please check the error log")
        session_log_data.append(
            [f'Duplicates found in data set: {duplicates_found}']
            )
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
        missing_values_removed_df = missing_values_removed_df.drop_duplicates(
            keep='first'
            )
        print("     Duplicates have been removed\n")
    else:
        print("     No duplicates found in the working data set.\n")
        no_duplicates_df = missing_values_removed_df

    print("Validating duplicates completed     <<<<<\n\n\n")

    return missing_values_removed_df


def validate_outliers(
        no_duplicates_df, session_log_data,
        atmos_outlier_log, atmos_outliers_url,
        wind_outlier_log, wind_outliers_url,
        wave_outlier_log, wave_outliers_url,
        temp_outlier_log, temp_outliers_url
        ):
    """
    Validates the dataframe for outliers and logs the process across various
    environmental metrics.

    This function checks the provided dataframe for outliers in
    atmospheric pressure, wind speed, wave characteristics, and temperature.
    It logs the validation process
    in the session log and updates the respective outlier logs if any
    outliers are found.

    Args:
    no_duplicates_df (pandas.DataFrame): The dataframe that has been
    cleaned of duplicates.
    session_log_data (list): The list used to log session activity.
    atmos_outlier_log (gspread.models.Worksheet): The Google Sheet
    worksheet for atmospheric outliers.
    atmos_outliers_url (str): The URL link to the atmospheric outliers log.
    wind_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
    for wind outliers.
    wind_outliers_url (str): The URL link to the wind outliers log.
    wave_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
    for wave outliers.
    wave_outliers_url (str): The URL link to the wave outliers log.
    temp_outlier_log (gspread.models.Worksheet): The Google Sheet
    worksheet for temperature outliers.
    temp_outliers_url (str): The URL link to the temperature outliers log.

    Returns:
    None: This function logs the detected outliers and updates the
    corresponding Google Sheets
    but does not return any value.
    """
    # Check for outliers
    print("Outlier Validation Started       <<<<<\n")
    session_log_data.append(['Outlier Validation Started'])
    session_log_data.append([str(pd.Timestamp.now())])

    numeric_df = no_duplicates_df.apply(
        lambda col: col.map(lambda x: pd.to_numeric(x, errors='coerce'))
        )

    atmospheric_outliers = check_for_outliers(
        numeric_df[['AtmosphericPressure']]
        )
    wind_outliers = check_for_outliers(numeric_df[['WindSpeed', 'Gust']])
    wave_outliers = check_for_outliers(
        numeric_df[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']]
        )
    temp_outliers = check_for_outliers(
        numeric_df[['AirTemperature', 'SeaTemperature']]
        )

    # Update Outlier Sheets
    if not atmospheric_outliers.empty:
        atmos_outlier_log.update(
            df_to_list_of_lists(atmospheric_outliers), 'A1'
            )
        print(
            "     Atmospheric Outliers Were Found:"
            " Check Atmos Outlier Log"
            )
        print(
            f"    \nThe link to the Atmospheric Outliers log is: \n\
            \n{atmos_outliers_url}\n\n"
            )
    if not wind_outliers.empty:
        wind_outlier_log.update(df_to_list_of_lists(wind_outliers), 'A1')
        print("     Wind Outliers Were Found: Check Atmos Outlier Log")
        print(
            f"    \nThe link to the Wind Outliers log is:\n\n"
            f"{wind_outliers_url}\n\n"
            )
    if not wave_outliers.empty:
        wave_outlier_log.update(df_to_list_of_lists(wave_outliers), 'A1')
        print(
            "     Wave Outliers Were Found:        Check Wave Outlier Log"
            )
        print(
            f"    \nThe link to the Wave Outliers log is:\n\n"
            f"{wave_outliers_url}\n\n"
            )
    if not temp_outliers.empty:
        temp_outlier_log.update(df_to_list_of_lists(temp_outliers), 'A1')
        print(
            "     Temp Outliers Were Found:        Check Temp  Outlier "
            "Log\n"
            )
        print(
            f"    \nThe link to the Temp Outliers log is:\n\n"
            f"{temp_outliers_url}\n\n"
            )

    print("Outlier Validation Completed     <<<<<\n\n\n")


def validate_date_format(
        master_df, session_log_data, error_log_data, date_time_url
        ):
    """
    Validates the date and time format in the dataframe and logs
    any inconsistencies.

    This function checks the 'time' column in the master dataframe
    for inconsistencies
    with the expected date-time format (YYYY-MM-DDTHH:MM:SSZ).
    It logs the validation process in the session log and records any
    errors found in the error log. Rows with incorrect date formats are
    removed, and the remaining data is standardized
    to a specific format.

    Args:
    master_df (pandas.DataFrame): The dataframe containing the master
    data to be validated.
    session_log_data (list): The list used to log session activity.
    error_log_data (list): The list used to log errors encountered
    during validation.
    date_time_url (str): The URL link to the date-time error log.

    Returns:
    tuple: A tuple containing the following:
        - pandas.DataFrame: The dataframe with validated and
          standardized date formats.
        - list: A list of logged date-time errors.
    """
    # Check for date inconsistencies
    date_time_error_log_data = []
    print("Date Validation Started       <<<<<\n")
    session_log_data.append(['Date and Time Validation Started'])
    session_log_data.append([str(pd.Timestamp.now())])
    date_time_pattern = r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
    inconsistent_date_format = master_df[
        ~master_df['time'].str.match(date_time_pattern)]

    if inconsistent_date_format.empty:
        print("     No Incorrect Date Formats Found")
        validated_data_df = master_df.copy()
        validated_data_df['time'] = pd.to_datetime(
            validated_data_df['time'], format='%Y-%m-%dT%H:%M:%SZ'
            )
        validated_data_df['time'] = validated_data_df['time'].dt.strftime(
            '%d-%m-%YT%H:%M:%S'
            )
    else:
        print("     Incorrect Date Formats Found\n")
        date_time_error_log_data.append(
            ['Inconsistent Date'
             ' and Time Formats']
            )
        date_time_error_log_data.append(
            inconsistent_date_format['time'].fillna('').astype(
                str
                ).values.tolist()
            )
        validated_data_df = master_df[
            master_df['time'].str.match(date_time_pattern)].copy()
        validated_data_df['time'] = pd.to_datetime(
            validated_data_df['time'], format='%Y-%m-%dT%H:%M:%SZ'
            )
        validated_data_df['time'] = validated_data_df['time'].dt.strftime(
            '%d-%m-%YT%H:%M:%S'
            )

    print("     Incorrect Date Formats Removed\n")
    # Log inconsistent date formats here
    # date_time_error_log.update(date_time_error_log_data, 'A1')
    print(
        f"    \nDate Inconsistancies found are written here: \n\
        \n{date_time_url}\n\n"
        )

    return validated_data_df, date_time_error_log_data


def update_all_logs(
        session_log_data, error_log_data,
        date_time_error_log_data, session_log,
        error_log, date_time_error_log
        ):
    """
    Updates all log sheets with the latest session, error, and
    date-time validation data.

    This function converts session logs, error logs, and date-time
    error logs from lists to a format suitable for Google Sheets,
    then updates the respective sheets with this data.
    It uses a helper function to handle the updates, ensuring each
    log is recorded accurately.

    Args:
    session_log_data (list): The list containing session log entries.
    error_log_data (list): The list containing error log entries.
    date_time_error_log_data (list): The list containing date-time
    error log entries.
    session_log (gspread.models.Worksheet): The Google Sheet
    worksheet for the session log.
    error_log (gspread.models.Worksheet): The Google Sheet
    worksheet for the error log.
    date_time_error_log (gspread.models.Worksheet): The Google Sheet
    worksheet for the date-time error log.

    Returns:
    None: This function updates the Google Sheets logs and does
    not return any value.
    """
    # Convert log lists to strings and update Google Sheets
    handle_log_update(
        session_log.update, session_log,
        df_to_list_of_lists(pd.DataFrame(session_log_data)),
        log_name='session log'
        )
    handle_log_update(
        error_log.update, error_log,
        df_to_list_of_lists(pd.DataFrame(error_log_data)),
        log_name='error log'
        )
    handle_log_update(
        date_time_error_log.update,
        date_time_error_log,
        df_to_list_of_lists(pd.DataFrame(date_time_error_log_data)),
        log_name='date time log'
        )


def df_to_list_of_lists(df):
    """
    Function to convert dataframe into list of lists.
    This is used by the outlier functions and the updater log functions
    to write the dataframes to the respective google sheet
    """
    return [df.columns.tolist()] + df.values.tolist()


def validate_master_data(
        master_data, session_log_data, error_log_data,
        session_log, error_log, date_time_error_log,
        validated_master_data, validated_master_data_url,
        atmos_outlier_log, atmos_outliers_url,
        wind_outlier_log, wind_outliers_url,
        wave_outlier_log, wave_outliers_url,
        temp_outlier_log, temp_outliers_url,
        date_time_url
        ):
    """
    Validates and cleans a marine data set through several checks
    and updates logs accordingly.

    This function performs a series of validation and cleaning
    steps on the input master data.
    The process includes:

    1. **Missing Values Validation:**
       - Identifies and logs rows with missing values.
       - Replaces missing values with NaN and removes affected rows
         from the dataset.

    2. **Duplicate Rows Validation:**
       - Detects duplicate rows, logs the issue, and removes them from
         the dataset.

    3. **Outlier Detection:**
       - Identifies outliers in numerical data columns such as
         AtmosphericPressure, WindSpeed, etc.
       - Logs detected outliers to specific logs for atmospheric, wind,
         wave, and temperature data.

    4. **Date Format Validation:**
       - Validates and corrects the ISO 8601 date and time format in
         the dataset.
       - Logs any rows with date format inconsistencies and attempts
        corrections.

    After the validations, the function updates session logs, error logs,
    and date-time logs in Google Sheets, and writes the cleaned, validated
    data back to a specified Google Sheet.

    Args:
    - master_data (list): The input dataset where the first
    row contains column headers.
    - session_log_data (list): A list to accumulate entries
      for the session log.
    - error_log_data (list): A list to accumulate entries for the error log.
    - session_log (gspread.models.Worksheet): The Google Sheet
      worksheet for the session log.
    - error_log (gspread.models.Worksheet): The Google Sheet
      worksheet for the error log.
    - date_time_error_log (gspread.models.Worksheet): The Google
      Sheet worksheet for the date-time error log.
    - validated_master_data (gspread.models.Worksheet): The Google
      Sheet where the validated master data will be written.
    - validated_master_data_url (str): The URL of the Google Sheet
      containing the validated master data.
    - atmos_outlier_log (gspread.models.Worksheet): The Google Sheet
      worksheet for atmospheric outlier data.
    - atmos_outliers_url (str): The URL of the Google Sheet for atmospheric
      outliers.
    - wind_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
      for wind outlier data.
    - wind_outliers_url (str): The URL of the Google Sheet for wind outliers.
    - wave_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
      for wave outlier data.
    - wave_outliers_url (str): The URL of the Google Sheet for wave outliers.
    - temp_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
      for temperature outlier data.
    - temp_outliers_url (str): The URL of the Google Sheet for temperature
      outliers.
    - date_time_url (str): The URL of the Google Sheet for date-time errors.

    Returns:
    - pd.DataFrame: The cleaned and validated master data as a
      pandas DataFrame.
    """
    # Initialise log lists
    date_time_error_log_data = []  # Initialize date_time_error_log_data here
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
        # validate data for missing values - output clean dataframe
        master_df = validate_missing_values(
            master_df, session_log_data,
            error_log_data
            )
        # validate data for duplicate and remove - output clean dataframe
        master_df = validate_duplicates(
            master_df, session_log_data,
            error_log_data
            )
        # check data for outliers and output to sheets if found
        validate_outliers(
            master_df, session_log_data,
            atmos_outlier_log, atmos_outliers_url,
            wind_outlier_log, wind_outliers_url,
            wave_outlier_log, wave_outliers_url,
            temp_outlier_log, temp_outliers_url
            )
        # validate data after date format errors and remove
        validated_data_df, date_time_error_log_data = validate_date_format(
            master_df, session_log_data, error_log_data, date_time_url
            )

        print("Date Validation Completed     <<<<<\n")
        session_log_data.append(['Data Validation Ended <<<<<<<<<<'])
        session_log_data.append([str(pd.Timestamp.now())])

        # Convert all elements in logs to strings
        session_log_data = \
            [[str(item) for item in sublist] for sublist in
             session_log_data]
        error_log_data = \
            [[str(item) for item in sublist] for sublist in
             error_log_data]
        date_time_error_log_data = \
            [[str(item) for item in sublist] for sublist in
             date_time_error_log_data]

    except Exception as e:
        session_log_data.append([f"An error occurred during validation: {e}"])
        print(f"An error occurred during validation: {e}")

    finally:
        update_all_logs(
            session_log_data, error_log_data,
            date_time_error_log_data, session_log, error_log,
            date_time_error_log
            )

    print("\n\n\n >>>>> Master Data Validation Completed <<<<<\n\n\n")
    print("Writing Validated Data To Google Sheets Started      <<<<<\n")
    set_with_dataframe(validated_master_data, validated_data_df)
    print("Writing Validated Data To Google Sheets Completed    <<<<<\n")
    print(
        f"    \nThe validated master data is written here:\n\n"
        f"{validated_master_data_url}\n\n"
        )
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
    validated_df['time'] = pd.to_datetime(
        validated_df['time'],
        format='%d-%m-%YT%H:%M:%S',
        errors='coerce', dayfirst=True
        )
    # Extract the date component only, ignoring the time
    validated_df['date_only'] = validated_df['time'].dt.strftime('%d-%m-%Y')


def validate_input_dates(
        date_str, reference,
        df_first_date, df_last_date
        ):
    """
    Validates an input date string to ensure it is in the correct format,
    represents a valid calendar date, and falls within
    the specified date range.

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
                "Date is not in the correct format or is invalid."
                )

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
                        "February in a leap year has only 29 days."
                        )
            else:
                if day > 28:
                    raise ValueError(
                        "February has only 28 days in a non-leap year."
                        )

        # Ensure the date falls within the available range
        if date < pd.to_datetime(df_first_date) or date > pd.to_datetime(
                df_last_date
                ):
            raise ValueError(
                f"Date {date_str} is outside the available data range."
                )
        return date
    except ValueError as ve:
        # Re-raise the error with a custom message
        raise ValueError(
            f"Invalid date: {ve}. Please enter the date in  'dd-mm-yyyy' "
            f"format."
            )


def get_user_dates(validated_df, error_log_data, error_log):
    """
    Prompts the user to select a start and end date from a validated dataset,
    checks the validity of the entered dates, and logs any errors.

    This function performs the following tasks:

    1. **Display Date Range:**
       - Shows the user the available date range in the master dataset.

    2. **User Input for Start Date:**
       - Prompts the user to enter a start date within the displayed range.
       - Validates the input date format and ensures it falls within the
         allowed range.
       - Logs any errors related to the start date.

    3. **User Input for End Date:**
       - Prompts the user to enter an end date within the displayed range.
       - Validates the input date format, ensures it falls within the
         allowed range,
         and that it is not earlier than the start date.
       - Logs any errors related to the end date.

    4. **Logging Errors:**
       - Writes any date-related errors to the error log in Google Sheets.

    Args:
    - validated_df (pd.DataFrame): A DataFrame containing a 'time' column
      with date and time data in the format '%d-%m-%YT%H:%M:%S'.
    - error_log_data (list): A list to accumulate entries for the error log.
    - error_log (gspread.models.Worksheet): The Google Sheet worksheet for
      the error log.

    Returns:
    - tuple: A tuple containing the validated start and end dates as strings in
      the format 'dd-mm-yyyy'.
    """
    # Declare variable for later use
    error_log_data = []

    # Convert the time column in validated_df to date time
    validated_df['time'] = pd.to_datetime(
        validated_df['time'],
        format='%d-%m-%YT%H:%M:%S'
        )

    # Get the first and last dates available in the data frame
    df_first_date = validated_df['time'].min().strftime('%d-%m-%Y')
    df_last_date = validated_df['time'].max().strftime('%d-%m-%Y')
    print(f"\n\nAvailable data range: {df_first_date} to {df_last_date}\n\n")
    print("You will be asked to enter a start date and and end date")
    print(">>> You can type 'quit' at any time to exit. <<<\n")

    # Prompt user for start date, until date is acceptable and formatted
    while True:
        user_input_start_date = input(
            f"Please enter the start date, or Type quit \n (Date in "
            f" the format 'dd-mm-yyyy' "
            f"\nbetween {df_first_date} and {df_last_date}): "
            )
        # Check to see if the user wants to quit
        if user_input_start_date.lower() == 'quit':
            print("Exiting program as requested.")
            exit()

        try:
            user_input_start_date = validate_input_dates(
                user_input_start_date,
                "start",
                df_first_date,
                df_last_date
                )
            break  # Exit loop if the date is valid
        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(
                ["Start Date Error", str(pd.Timestamp.now())]
                )
            error_log_data.append(["Start Date Error", str(e)])
            print("Start Date Error:\n")
            print(f"You Entered: {user_input_start_date}  <<<<<\n")
            print(
                f"\n A detailed description of the error\n has been  "
                f"appended to the error log."
                )
            print("\nPlease enter the date in 'dd-mm-yyyy' format.\n\n")

    # Prompt user for end date
    while True:
        user_input_end_date = input(
            f"\n\nPlease enter the end date, or Type quit \n (Date in the "
            f"format 'dd-mm-yyyy' \n  "
            f"within {df_first_date} and {df_last_date}): "
            )
        # Check to see if the user wants to quit
        if user_input_end_date.lower() == 'quit':
            print("Exiting program as requested.")
            exit()

        try:
            user_input_end_date = validate_input_dates(
                user_input_end_date,
                "end",
                df_first_date,
                df_last_date
                )
            if user_input_end_date < user_input_start_date:
                raise ValueError(
                    "The end date cannot be earlier than the start date."
                    )
            break  # Exit loop if the date is valid
        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(["End Date Error", str(pd.Timestamp.now())])
            error_log_data.append(["End Date Error", str(e)])
            print(f"End Date Error:\n")
            print(f"You Entered: {user_input_end_date}  <<<<<\n")
            print(
                f"\n A detailed description of the error\n has been appended "
                f"to the error log."
                )
            print("\nPlease enter the date in 'dd-mm-yyyy' format.\n\n")

    # Convert validated start and end dates to string format
    # for further processing
    user_input_start_date_str = user_input_start_date.strftime('%d-%m-%Y')
    user_input_end_date_str = user_input_end_date.strftime('%d-%m-%Y')

    # Write any errors to log
    handle_log_update(
        error_log.update, error_log,
        df_to_list_of_lists(pd.DataFrame(error_log_data)),
        log_name='error log'
        )

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
        (pd.to_datetime(
            validated_df['date_only'],
            format='%d-%m-%Y'
            ) >= start_date) &
        (pd.to_datetime(
            validated_df['date_only'],
            format='%d-%m-%Y'
            ) <= end_date)
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
        working_data_df['time']
        ).dt.strftime('%d-%m-%Y %H:%M:%S')

    return working_data_df


def get_data_selection(error_log):
    """
    Prompts the user to select specific data columns to display from a
    predefined list of options.

    The function provides a menu with the following choices:
    1. All Data
    2. Atmospheric Pressure
    3. Wind Speed and Gust
    4. Wave Height, Wave Period, and Mean Wave Direction
    5. Air Temperature and Sea Temperature
    6. Exit Output Options

    The user is prompted to enter a number corresponding to their choice.
    The function:
    - Validates the user input to ensure it is a valid option.
    - Logs any errors if the user inputs an invalid selection.
    - Returns a list of column names based on the user's choice,
      or an empty list if the user chooses to exit.

    The function will continue prompting the user until a valid
    selection is made or the user opts to exit.

    Args:
    - error_log (gspread.models.Worksheet): The Google Sheet
      worksheet for the error log.

    Returns:
    - list: A list of column names to be displayed based on user selection.
      Returns an empty list if the user chooses to exit.
    """
    # Initialise selection storage
    selected_columns = []
    error_log_data = []
    while True:
        # Output Selection Options
        print("\n\n\nSelect the data you want to display:")
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
                    "1 and 6."
                    )

            # If a valid selection is made, break out of the loop and
            # return the selected columns
            break

        except ValueError as e:
            # Append error details to the error log and inform the user
            error_log_data.append(
                ["Data Set Selection Error", str(pd.Timestamp.now())]
                )
            error_log_data.append(["You Input ", selection])
            error_log_data.append(["Error Description", str(e)])
            print(f"Data Set Selection Error:\n")
            print(f"You Entered: {selection}    <<<<<\n")
            print(
                f"Invalid data selection input.\n A detailed description "
                f"of the error\n has been appended to the error log."
                )
            print("\nPlease enter 1, 2, 3, 4, 5, 6 to exit\n")

        # Write any errors to log
        handle_log_update(
            error_log.update, error_log,
            df_to_list_of_lists(pd.DataFrame(error_log_data)),
            log_name='error log'
            )

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
            f"more rows not displayed."
            )

    return allow_screen, allow_graph, allow_sheet


def get_output_selection(
        user_output_df, user_data_output, selected_columns, allow_screen,
        allow_graph, allow_sheet, num_rows, SCOPED_CREDS,
        user_data_output_url, error_log, graphical_output_data_url
        ):
    """
    Prompts the user to select an output option for a DataFrame and performs
    the corresponding action.

    This function provides an interactive loop that allows the user to choose
    how they want to handle the output of the provided
    DataFrame (`user_output_df`).
    The options include displaying data on the screen, plotting a graph,
    writing data
    to a Google Sheet, or exiting the loop.

    The function handles the following actions based on user selection:
    1. Displaying the DataFrame on the screen (if `allow_screen` is True).
    2. Generating and displaying a graph of the
       DataFrame (if `allow_graph` is True).
    3. Writing the DataFrame to a Google Sheet (if `allow_sheet` is True).
    4. Exiting the loop.

    Args:
    - user_output_df (pd.DataFrame): The DataFrame to be processed and
      outputted based on user selection.
    - selected_columns (list of str): The list of column names available
    in `user_output_df` for graphing or other operations.
    - allow_screen (bool): Flag indicating whether the option to display data
      on the screen is available.
    - allow_graph (bool): Flag indicating whether the option to generate a
      graph is available.
    - allow_sheet (bool): Flag indicating whether the option to write data
      to a Google Sheet is available.
    - num_rows (int): Number of rows in `user_output_df`. Used to determine
      how many rows to display on the screen.
    - SCOPED_CREDS (Credentials): Credentials for accessing Google Sheets.
    - user_data_output_url (str): URL of the Google Sheet where the data
      will be written.
    - error_log (gspread.models.Worksheet): The Google Sheet worksheet for
      the error log.

    Returns:
    - None: The function handles user interactions and performs actions
      based on user selection.
    """

    # Inner loop allowing user select different output options
    while True:
        # Get the action from user with validation
        output_selection = get_valid_data_output_selection(
            allow_screen,
            allow_graph,
            allow_sheet,
            error_log
            )
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
                user_requested_graph(
                    user_output_df, x_col, y_cols, title, SCOPED_CREDS,
                    graphical_output_data_url
                    )
            # If user selects 3 - output to google sheet
            elif output_selection == 3:
                # Option 3 Write Data To Google Sheet
                set_with_dataframe(user_data_output, user_output_df)
                print("\nData Written To Google Sheet")
                print(
                    f"    \nData Output Can Be Found Here:"
                    f"\n\n{user_data_output_url}\n\n"
                    )
            # If user select 4 - loop ends
            elif output_selection == 4:
                # Option 4 Exit the loop
                print("\nExited Output Options ......\n")
                break

        except ValueError as e:
            print("Error in output selection")


def data_initialisation_and_validation(
        session_log_url, gael_force_error_log_url,
        session_log, error_log, unvalidated_master_data,
        date_time_error_log, atmos_outliers_url,
        validated_master_data, validated_master_data_url,
        wave_outlier_log, temp_outlier_log,
        atmos_outlier_log, wind_outliers_url,
        wave_outliers_url, temp_outliers_url,
        wind_outlier_log, date_time_url
        ):
    """
    Initializes and validates marine data, and sets up the validated data
    for the session.

    This function performs the following steps:
    1. **Load Marine Data:**
       - Loads the marine data from the specified input sheet.
       - Retrieves session and error logs for the current session.

    2. **Validate Data:**
       - Executes a series of validation checks on the loaded marine data.
       - Cleans the data by removing missing values, duplicates, and outliers.
       - Validates and corrects date and time formats.
       - Logs the results of these validations in Google Sheets.

    3. **Store Validated Data:**
       - Stores the validated data in a specified location for further use
         in the session.

    This function is intended to be run once per session to ensure data
    integrity and prepare the data for subsequent processing.

    Args:
    - session_log_url (str): URL for the Google Sheet containing the
      session log.
    - gael_force_error_log_url (str): URL for the Google Sheet containing
      the error log.
    - session_log (gspread.models.Worksheet): The Google Sheet worksheet for
      the session log.
    - error_log (gspread.models.Worksheet): The Google Sheet worksheet for
      the error log.
    - unvalidated_master_data (list): The raw marine data to be validated.
    - date_time_error_log (gspread.models.Worksheet): The Google Sheet
      worksheet for date and time errors.
    - atmos_outliers_url (str): URL for the Google Sheet containing
      atmospheric outliers.
    - validated_master_data (gspread.models.Worksheet): The Google Sheet
      worksheet for validated master data.
    - validated_master_data_url (str): URL where the validated master data
      will be written.
    - wave_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
      for wave outliers.
    - temp_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
      for temperature outliers.
    - atmos_outlier_log (gspread.models.Worksheet): The Google Sheet worksheet
      for atmospheric outliers.
    - wind_outliers_url (str): URL for the Google Sheet containing
      wind outliers.
    - wave_outliers_url (str): URL for the Google Sheet containing
      wave outliers.
    - temp_outliers_url (str): URL for the Google Sheet containing
      temperature outliers.
    - wind_outlier_log (gspread.models.Worksheet): The Google Sheet
      worksheet for wind outliers.
    - date_time_url (str): URL for the Google Sheet containing date and
      time validation errors.

    Returns:
    - pd.DataFrame: The validated data frame ready for use in the session.
    """
    print("\n")
    print("#############################################################")
    print(">>>>>   Phase 1. Starting Data Validation Process       <<<<<")
    print("#############################################################")
    print("\n")
    # Load the marine data for validation
    master_data, session_log_data, error_log_data = (
        load_marine_data_input_sheet(
            session_log_url, gael_force_error_log_url,
            session_log, error_log, unvalidated_master_data
            ))

    # Create a validated data frame for use in the app
    validated_df = validate_master_data(
        master_data, session_log_data, error_log_data,
        session_log, error_log, date_time_error_log,
        validated_master_data, validated_master_data_url,
        atmos_outlier_log, atmos_outliers_url,
        wind_outlier_log, wind_outliers_url,
        wave_outlier_log, wave_outliers_url,
        temp_outlier_log, temp_outliers_url,
        date_time_url
        )

    return validated_df


def convert_dataframe(df, x_col, y_cols):
    """
    Convert specific columns in a DataFrame to appropriate data types.

    This function converts the data types of specified columns in a DataFrame
    to ensure that they are in the correct format for analysis or plotting.

    The function performs the following conversions:
    1. **Datetime Conversion:**
       - Converts the column specified by `x_col` to datetime format.
       - Handles invalid parsing by setting errors to 'coerce', which
         results in NaT (Not a Time) for invalid entries.

    2. **Numeric Conversion:**
       - Converts the columns specified in `y_cols` to numeric format.
       - Handles invalid parsing by setting errors to 'coerce', which
         results in NaN (Not a Number) for invalid entries.

    Args:
    - df (pd.DataFrame): The DataFrame containing the data to be converted.
    - x_col (str): The name of the column to be converted to datetime format.
    - y_cols (list of str): A list of column names to be converted to
      numeric format.

    Returns:
    - pd.DataFrame: The DataFrame with specified columns converted to
      the appropriate data types.
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

    This function uses the Google Sheets API to write data to a specific sheet
    within a Google Spreadsheet. The data is written starting from cell A1.

    Args:
    - service: Authorized Google Sheets API service instance. This is an
      instance
      of the `googleapiclient.discovery.Resource` class, which allows
     interaction with the Google Sheets API.
    - spreadsheet_id (str): The unique identifier of the Google
      Spreadsheet where the data will be written. This can be found in the
      URL of the spreadsheet.
    - sheet_name (str): The name of the sheet within the spreadsheet
      where the data will be written. For example, 'Sheet1'.
    - values (list of lists): The data to be written to the sheet.
      Each inner list represents a row of data, and each element within the
      inner list represents a cell in that row.

    Returns:
    - None: This function does not return a value. It writes the
      data directly to the specified Google Sheet.

    Raises:
    - HttpError: If there is an issue with the HTTP request to the
      Google Sheets API.
    - GoogleAuthError: If there is an issue with authentication
      while accessing the API.
    - Exception: For any other unexpected errors that occur during
      the operation.
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

    This function uses the Google Sheets API to identify and remove any charts
    that are embedded in the specified Google Spreadsheet. It performs a batch
    update request to delete all found charts.

    Args:
    - service: Authorized Google Sheets API service instance.
      This is an instance of the `googleapiclient.discovery.Resource` class,
      which allows interaction with the Google Sheets API.
    - spreadsheet_id (str): The unique identifier of the Google
      Spreadsheet where the charts will be deleted. This ID can be found in
      the URL of the spreadsheet.

    Returns:
    - None: This function does not return a value.
      It directly performs the deletion of charts in the specified Google
      Sheet.

    Raises:
    - HttpError: If there is an issue with the HTTP request to the
      Google Sheets API.
    - GoogleAuthError: If there is an issue with authentication while
      accessing the API.
    - Exception: For any other unexpected errors that occur during
      the operation.
    """
    try:
        # Retrieve the spreadsheet information
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
            ).execute()

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
                f"the Google Sheet."
                )
        else:
            print("No charts found to delete.")

    except HttpError as e:
        print(f"HTTP error occurred: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def add_chart_to_sheet(
        service, spreadsheet_id, sheet_id, x_col, y_cols,
        title, graphical_output_data_url
        ):
    """
    Add a chart to the specified Google Sheet with the legend positioned
    on the right, and configure the chart to use the first row as headers.

    This function first deletes any existing charts in the specified Google
    Sheet to ensure that the new chart is added cleanly. It then creates a
    line chart where the x-axis is defined by the `x_col` and multiple
    y-axis series are defined by `y_cols`. The chart is placed on the sheet
    with a title and the legend positioned on the right.

    Args:
    - service: Authorized Google Sheets API service instance. This instance
      is used to interact with the Google Sheets API.
    - spreadsheet_id (str): The unique identifier of the Google Spreadsheet
      where the chart will be added. This ID can be found in the URL of the
      spreadsheet.
    - sheet_id (int): The ID of the sheet within the spreadsheet where the
      chart will be added.
    - x_col (str): The header name or index of the column to be used for
      the x-axis of the chart.
    - y_cols (list of str): List of header names or indices of the columns
      to be used for the y-axis (multiple series) of the chart.
    - title (str): The title of the chart.

    Returns:
    - None: This function does not return a value. It directly performs
      the addition of the chart to the specified Google Sheet.

    Raises:
    - HttpError: If there is an issue with the HTTP request to the Google
      Sheets API.
    - GoogleAuthError: If there is an issue with authentication while
      accessing the API.
    - Exception: For any other unexpected errors that occur during the
      operation.
    """
    delete_existing_charts(service, spreadsheet_id)
    try:
        # Determine the data range for the chart
        # Assuming that the number of rows matches the length of y_cols
        end_row_index = len(
            y_cols
            ) + 1

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
                                    {
                                        'position': 'BOTTOM_AXIS',
                                        'title': x_col
                                        },  # X-axis title
                                    {
                                        'position': 'LEFT_AXIS',
                                        'title': 'Values'
                                        }  # Y-axis title
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
            f"\n\n{graphical_output_data_url}\n\n"
            )
    except HttpError as e:
        print(f"HTTP error occurred while creating the chart: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred while creating the chart: {e}")


def user_requested_graph(df, x_col, y_cols, title, SCOPED_CREDS, graphical_output_data_url):
    """
    Writes data from a Pandas DataFrame to a specified Google Sheet and
    creates a chart.

    This function takes a Pandas DataFrame, extracts specified columns,
    and writes the
    data to a Google Sheet. It then creates a chart in the Google Sheet
    using the+ specified columns.

    Args:
    - df (pd.DataFrame): The DataFrame containing the data to be written to
      the Google Sheet.
    - x_col (str): The name of the column in `df` to be used as the x-axis
      in the chart.
    - y_cols (list of str): A list of column names in `df` to be used as the
      y-axis in the chart.
    - title (str): The title of the chart to be created in the Google Sheet.
    - SCOPED_CREDS: Authorized Google Sheets API credentials instance.

    Returns:
    - None: This function does not return a value. It performs data writing
      and chart creation
      operations directly in the specified Google Sheet.

    Raises:
    - HttpError: If there is an issue with the HTTP request to the
      Google Sheets API.
    - GoogleAuthError: If there is an issue with authentication while
      accessing the API.
    - Exception: For any other unexpected errors that occur during
      the operation.
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
                                                     y_cols]
                )

        # Write data to the sheet
        write_data_to_sheet(service, spreadsheet_id, sheet_name, values)

        # Add a chart to the sheet
        add_chart_to_sheet(
            service, spreadsheet_id, sheet_id, x_col, y_cols,
            title, graphical_output_data_url
            )

    except HttpError as e:
        print(f"HTTP error occurred: {e}")
    except GoogleAuthError as e:
        print(f"Authentication error occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def get_valid_data_output_selection(
        allow_screen, allow_graph, allow_sheet, error_log
        ):
    """
    Prompts the user to select an output option and ensures that only valid
    selections are allowed.

    This function provides an interactive menu for the user to choose how
    they want to handle the output
    of the selected data. The available options depend on the flags provided
    (allow_screen, allow_graph, allow_sheet), which control whether each
    option should be available. It ensures that the input is valid
    and logs any errors if the user makes an invalid selection.

    Args:
    - allow_screen (bool): Flag indicating if the option to print data to
      the screen is available.
    - allow_graph (bool): Flag indicating if the option to create a graph
      is available.
    - allow_sheet (bool): Flag indicating if the option to write data to a
      Google Sheet is available.
    - error_log: Function or object for updating the error log with any
      issues encountered.

    Returns:
    - int: The user's valid output selection, which is an integer
      corresponding to their choice:
      1 for "Print to Screen", 2 for "Create Graph", 3 for "Write to
      Google Sheet", or 4 for "Exit".

    Raises:
    - ValueError: If the user input cannot be converted to an integer or is
      not a valid selection.
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
            "\nEnter the number corresponding to your desired output: "
            )

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
                ["Output Selection Error", str(pd.Timestamp.now())]
                )
            error_log_data.append(["Output Selection Error", str(e)])
            print(f"Output Selection Error:\n")
            print(f"You Entered: {user_input}  <<<<<\n")
            print(
                f"\nA detailed description of the error\nhas been appended "
                f"to the error log."
                )
            print("Invalid selection. Please enter a number between 1 and 4.")

        # Write any errors to log
        handle_log_update(
            error_log.update, error_log,
            df_to_list_of_lists(pd.DataFrame(error_log_data)),
            log_name='error log'
            )

    return output_selection


def get_continue_yn(error_log):
    """
    Introduces the application to the user and prompts them to decide whether
    they want to continue.

    This function displays an introduction message about the application,
    which includes information on:
    - Data Validation
    - Data Interrogation

    After displaying the introduction, it enters a loop asking the user if
    they want to continue. The user should respond with 'y' to continue or
    'n' to exit. If the user inputs 'y', the function returns 'y'
    and the application continues. If 'n' is input, the application exits.
    Invalid inputs are handled by displaying an error message, logging the
    error details, and repeating the prompt.

    Args:
    - error_log: Function or object for updating the error log with any
      issues encountered.

    Returns:
    - str: The user's response, which will be 'y' to continue or 'n' to exit.

    Raises:
    - ValueError: If the user input is not 'y' or 'n'.
    """
    error_log_data = []
    # Display app introduction
    print("#######################################################")
    print("Welcome to the Weather Data Analysis Application.\n")
    print(
        "This application has two parts:\n - Data Validation\n - "
        "Data Interrogation.\n"
        )
    print("#######################################################")

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
                ["User Input Error", str(pd.Timestamp.now())]
                )
            error_log_data.append(["You Entered", str(proceed)])
            error_log_data.append(["User Input Error", str(e)])
            print(f"User Input Error:\n")
            print(f" Error: {e}\n")
            print(
                f"Invalid input.>>> {proceed} <<<\n\nA detailed description "
                f"of the error\n has been appended to the error log."
                )
            print("\nPlease enter (n) to exit\n")

        # Write any errors to log
        handle_log_update(
            error_log.update, error_log,
            df_to_list_of_lists(pd.DataFrame(error_log_data)),
            log_name='error log'
            )


def main():
    """
    Main function that controls the overall functionality of the application.

    This function orchestrates the workflow of the application, including:
    1. Prompting the user to continue or exit the application.
    2. Initializing and validating the master data from Google Sheets.
    3. Checking if the data validation was successful and proceeding if valid.
    4. Formatting the validated data frame's date columns.
    5. Allowing the user to specify a date range for filtering the data.
    6. Filtering the data frame based on the user-specified date range.
    7. Formatting the filtered data frame for display purposes.
    8. Providing options for users to select specific data columns for output.
    9. Determining and managing output options based on the number of rows
       in the data.
    10. Generating output based on user preferences and choices.
    11. Logging and writing error data to an error log sheet if applicable.

    This function performs the following steps:
    - Initializes the Google Sheets connection and retrieves required data.
    - Displays an introduction message and prompts the user to continue.
    - Loads and validates the master data.
    - If validation is successful, proceeds to filter and format the data
      based on user input.
    - Allows the user to select data columns and choose output options such
      as displaying on screen,
      creating a graph, or writing to a Google Sheet.
    - Handles errors and writes them to an error log if necessary.

    Logs:
    - Updates the error log with details of any errors encountered during
      execution.

    Exit Conditions:
    - Exits if the data is not validated or if the user inputs invalid dates
      or choices.

    Returns:
    - None
    """
    error_log_data = []
    # Initialise Sheets On Load
    sheet_initialisation = initialise_google_sheets()
    if sheet_initialisation:
        (SHEET, SCOPED_CREDS, unvalidated_master_data, validated_master_data,
         user_data_output, session_log, error_log,
         date_time_error_log, graphical_output_sheet,
         atmos_outlier_log, wind_outlier_log, wave_outlier_log,
         temp_outlier_log, google_worksheet, marine_master_data_url,
         validated_master_data_url, user_data_output_url, session_log_url,
         gael_force_error_log_url, atmos_outliers_url, wind_outliers_url,
         wave_outliers_url, temp_outliers_url, date_time_url,
         graphical_output_data_url) = sheet_initialisation

    while True:
        # Introduce the app and ask user if they want to continue
        decision = get_continue_yn(error_log)

        # Initialise the sheets and validate the data
        validated_df = data_initialisation_and_validation(
            session_log_url, gael_force_error_log_url,
            session_log, error_log, unvalidated_master_data,
            date_time_error_log, atmos_outliers_url,
            validated_master_data, validated_master_data_url,
            wave_outlier_log, temp_outlier_log,
            atmos_outlier_log, wind_outliers_url,
            wave_outliers_url, temp_outliers_url,
            wind_outlier_log, date_time_url
            )

        # Check to see if the data has been validated
        if validated_df is None:
            print(
                "Error: Data has not been validated\n The application cannot "
                "continue..... \nBye..."
                )
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
                f"\n\n{gael_force_error_log_url}\n\n"
                )
        # Convert the data frame date format to dd-mm-yyyy
        format_df_date(validated_df)

        # Outer Loop - get dates from user for specified range
        while True:
            # Get dates from user to interrogate validated data frame
            user_input_start_date_str, user_input_end_date_str = (
                get_user_dates(validated_df, error_log_data, error_log))

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
            date_filtered_df = filter_data_by_date(
                validated_df,
                user_input_start_date,
                user_input_end_date
                )

            # Format the dataframe for display, converting date format
            # to dd-mm-yyyy
            working_data_df = format_df_data_for_display(date_filtered_df)
            # Middle loop - getting specific data set for user output
            while True:
                # Get users selection for data output columns
                selected_columns = get_data_selection(error_log)
                if not selected_columns:
                    break
                user_output_df = working_data_df[selected_columns]
                num_rows = len(user_output_df)
                print(f"\nThere are {num_rows} rows of data.     <<<<<\n")

                # Determine output options based on rows
                allow_screen, allow_graph, allow_sheet = (
                    determine_output_options(num_rows))
                # Create output based on user selection
                get_output_selection(
                    user_output_df, user_data_output, selected_columns,
                    allow_screen, allow_graph, allow_sheet,
                    num_rows, SCOPED_CREDS, user_data_output_url,
                    error_log, graphical_output_data_url
                    )

    # Write error log to error log sheet if there are errors to be written
    if error_log_data:
        error_log_data = [[str(item) for item in sublist] for sublist in
                          error_log_data]
        print(f"The Error Log Contains:\n\n {error_log_data}")
        handle_log_update(
            error_log.update, error_log,
            df_to_list_of_lists(pd.DataFrame(error_log_data)),
            log_name='error log'
            )


main()
