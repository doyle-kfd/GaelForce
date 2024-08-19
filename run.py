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
date_time_log =  SHEET.worksheet('date_time')                                    # Date Time format error log

# define the sheets to be user for outlier output
atmos_outlier_log = SHEET.worksheet('atmos_outliers')
wind_outlier_log =  SHEET.worksheet('wind_outliers')
wave_outlier_log =  SHEET.worksheet('wave_outliers')
temp_outlier_log =  SHEET.worksheet('temp_outliers')
       



# Initialise the sheets for use
validated_master_data.clear()
user_data_report.clear()  
session_log.clear()
error_log.clear()
atmos_outlier_log.clear()
wind_outlier_log.clear()
wave_outlier_log.clear()
temp_outlier_log.clear()
date_time_log.clear()                                                              
                                                       


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
    z_scores = zscore(df)                            # calculate a z score for all values in dataframe
    abs_z_scores = abs(z_scores)                     # take the absolute value 
    outliers = (abs_z_scores > 3).any(axis=1)        # creates a bolean based on whether the absolute number is  > 3
    return df[outliers]                              # return the dataframe of the outliers identified


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
            
            ######session_log.update([['We found missing values in the master data, please check the error log    <<<<<<<']], 'A10')
            ######session_log.update([['{}'.format(log_timestamp)]], 'F10')
            ######error_log.update([['We found missing values in the master data, in the following columns:']], 'A3')
            ######error_log.update([['{}'.format(log_timestamp)]], 'F3')
            print(missing_values)
            ######error_log.update([['{}'.format(missing_values)]], 'A4')

            # Convert count of missing values series to a dataFrame for formatting to sheet
            missing_values_df = missing_values[missing_values > 0].reset_index()
            missing_values_df.columns = ['Column', 'Missing Values']
            missing_values.columns = ['Column', 'Missing Values']

            # Write the DataFrame to the error log
            for i, row in missing_values_df.iterrows():
                ######error_log.update([[row['Column'], row['Missing Values']]], f'A{i+4}') 
                print("\n")                                                               # Print them to the screen
        else:
            print("There were no cells with missing data")                            # Print to screen message no errors
            print(missing_values)
            #######session_log.update([['We found no missing values in the master data']], 'A10')
            print("\n")

        # Remove any rows that have no values
        if missing_values.any():
            ######session_log.update([['Removing data with no values from master data']], 'A12')
            ######log_timestamp = pd.Timestamp.now()
            ######session_log.update([['{}'.format(log_timestamp)]], 'f12')
            print("Starting to clean data frame\n")
            missing_values_removed_df = master_df.dropna()                                                 # Create a new df having dropped NaN values
            ######session_log.update([['Finished removing data with no values from master data']], 'A13')
            log_timestamp = pd.Timestamp.now()
            ######session_log.update([['{}'.format(log_timestamp)]], 'F13')
            missing_values = missing_values_removed_df.isnull().sum()                                      # count the missing values
            print("The Result is")
            print(missing_values)
            print("Missing value validation completed\n")
            print("\n\n\n")
            print(missing_values_removed_df)
            print("\n\n\n")
        #
        #
        # Check for duplicate rows
        #
        #
        print("Validating duplicates started\n")
        duplicates_found = missing_values_removed_df.duplicated(keep=False).sum()                           # Get a count of duplicates found
                  
        if duplicates_found.any():                                                                # If there are any duplicates found
            print(f"There are {duplicates_found} duplicates in the working data set\n")
            #####session_log.update([['Duplicates found in data set: {} , please check the error log    <<<<<<<<'.format(duplicates_found)]], 'A15') # Write the number of ducplicates to the session log
            log_timestamp = pd.Timestamp.now()
            ######session_log.update([['{}'.format(log_timestamp)]], 'F15')
            print("\n\n\n Duplicates Found")


            # Retrieve duplicate rows and store
            ######error_log.update([['Duplicate Rows Found']], 'A15')
            log_timestamp = pd.Timestamp.now()
            ######error_log.update([['{}'.format(log_timestamp)]], 'F15')

            # Filter the DataFrame to get the duplicate rows
            duplicates_df = missing_values_removed_df[missing_values_removed_df.duplicated(keep=False)]

            # Output starting at row 17
            start_row = 17
            for i, row in duplicates_df.iterrows():
                print(i)
                print(row)

                # Convert the row to a list and format it as needed for your error_log
                row_data = row.tolist()
                ###### error_log.update([row_data], f'A{start_row}')  # Update error log with the row data

                # Move to the next row in the error log
                start_row += 1

            # Create a new data frame with no duplicates.
            no_duplicates_df = missing_values_removed_df.drop_duplicates(keep='first')

        else:
            print("No duplicates found in the working data set.\n")

        print("Validating duplicates ended\n")

        #
        #
        # Check for outliers
        #
        #
        ###### session_log.update([['Outlier Validation Started']], 'A17')
        log_timestamp = pd.Timestamp.now()
        ###### session_log.update([['{}'.format(log_timestamp)]], 'F17')

        print("\n\n\nStarting Outliers Validation\n\n\n")

        # Convert cell values to integers for mathamtical use
        numeric_df = no_duplicates_df.apply(lambda col: col.map(lambda x: pd.to_numeric(x, errors='coerce')))

        # Create dataframes for use in outlier validation
        atmospheric_specific_df = numeric_df[['AtmosphericPressure']]
        wind_specific_df = numeric_df[['WindSpeed', 'Gust']]
        wave_specific_df = numeric_df[['WaveHeight', 'WavePeriod', 'MeanWaveDirection']]
        temp_specific_df = numeric_df[['AirTemperature', 'SeaTemperature']]

        # Create outliers from specific data frames
        atmospheric_outliers = check_for_outliers(atmospheric_specific_df)
        wind_outliers = check_for_outliers(wind_specific_df)
        wave_outliers = check_for_outliers(wave_specific_df)
        temp_outliers = check_for_outliers(temp_specific_df)




        
        # Start the atmospheric outlier validation process
        atmos_outlier_log.update([['Atmospheric Outliers Processing Started']], 'A1')
        log_timestamp = pd.Timestamp.now()
        ###### atmos_outlier_log.update([['{}'.format(log_timestamp)]], 'F1')
        ###### atmos_outlier_log.update([['Atmospheric Pressure']], 'A3')

        # Initilise starting positions for atmospheric outliers in google sheets
        start_row = 4
        max_rows_to_display = 10

        # process each atmospheric outlier, writing any entries to the atmos_output_log
        for i, row in atmospheric_specific_df.iterrows():
            if i >= max_rows_to_display:
                break
            # Convert the row to a list and format it as needed for your error_log
            row_data = row.tolist()
            ###### atmos_outlier_log.update([row_data], f'A{start_row}')  # Update error log with the row data


            print(f"Processing row {i+1} with data: {row_data}")

            # Move to the next row in the error log
            start_row += 1

        print("Atmospheric Outliers:\n", atmospheric_outliers)

        ###### wind_outlier_log.update([['Wind Outliers Processing Started']], 'A1')
        log_timestamp = pd.Timestamp.now()
        ###### wind_outlier_log.update([['{}'.format(log_timestamp)]], 'F1')
        ###### wind_outlier_log.update([['Wind Speed']], 'A3')
        ######wind_outlier_log.update([['Wind Gusts']], 'B3')
        # Start the wind specific outlier validation process
        # Initilise starting positions for outliers in google sheets
        start_row = 4
        max_rows_to_display = 10

        for i, row in wind_specific_df.iterrows():
            if i >= max_rows_to_display:
                break
            # Convert the row to a list and format it as needed for your error_log
            row_data = row.tolist()
            ###### wind_outlier_log.update([row_data], f'A{start_row}')  # Update error log with the row data


            print(f"Processing row {i+1} with data: {row_data}")

            # Move to the next row in the error log
            start_row += 1

        print("Wind Outliers:\n", wind_outliers)

        ###### wave_outlier_log.update([['Wave Outliers Processing Started']], 'A1')
        log_timestamp = pd.Timestamp.now()
        ####### wave_outlier_log.update([['{}'.format(log_timestamp)]], 'F1')
        ###### wave_outlier_log.update([['Wave Height']], 'A3')
        ###### wave_outlier_log.update([['Wave Period']], 'B3')
        ###### wave_outlier_log.update([['Wave Mean']], 'C3')
        # Start the wave specific outlier validation process
        # Initilise starting positions for outliers in google sheets
        start_row = 4
        max_rows_to_display = 10

        for i, row in wave_specific_df.iterrows():
            if i >= max_rows_to_display:
                break
            # Convert the row to a list and format it as needed for your error_log
            row_data = row.tolist()
            ###### wave_outlier_log.update([row_data], f'A{start_row}')  # Update error log with the row data


            print(f"Processing row {i+1} with data: {row_data}")

            # Move to the next row in the error log
            start_row += 1

        print("Wave Outliers:\n", wave_outliers)
        
        ###### temp_outlier_log.update([['Temperature  Outliers Processing Started']], 'A1')
        log_timestamp = pd.Timestamp.now()
        ###### temp_outlier_log.update([['{}'.format(log_timestamp)]], 'F1')
        ###### temp_outlier_log.update([['AirTemperature']], 'A3')
        ###### temp_outlier_log.update([['SeaTemperature']], 'B3')
        # Start the wave specific outlier validation process
        # Initilise starting positions for outliers in google sheets
        start_row = 4
        max_rows_to_display = 10

        for i, row in temp_specific_df.iterrows():
            if i >= max_rows_to_display:
                break
            # Convert the row to a list and format it as needed for your error_log
            row_data = row.tolist()
            ###### temp_outlier_log.update([row_data], f'A{start_row}')  # Update error log with the row data


            print(f"Processing row {i+1} with data: {row_data}")

            # Move to the next row in the error log
            start_row += 1

        print("Temperature Outliers:\n", temp_outliers) 
        

        print("\n\n\nEnded Outliers Validation\n\n\n")
        ###### session_log.update([['Outlier Validation Finished']], 'A18')
        log_timestamp = pd.Timestamp.now()
        ###### session_log.update([['{}'.format(log_timestamp)]], 'F18')




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

        session_log.update([['Date and time validation started']], 'A30')
        log_timestamp = pd.Timestamp.now()
        session_log.update([['{}'.format(log_timestamp)]], 'F30')
        date_time_log.update([['Date and time validation started']], 'A1')
        date_time_log.update([['{}'.format(log_timestamp)]], 'F1')

        pattern = r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'


        
        inconsistent_date_format = no_duplicates_df[~no_duplicates_df['time'].str.match(pattern)]

        start_row = 4
        max_rows_to_display = 10

        if inconsistent_date_format.empty:
            print("All time values are in the correct format.")
    
            # Convert and reformat the time column
            no_duplicates_df['time'] = pd.to_datetime(no_duplicates_df['time'], format='%Y-%m-%dT%H:%M:%SZ')
            no_duplicates_df['time'] = no_duplicates_df['time'].dt.strftime('%d-%m-%y')
    
            print("Time values have been reformatted to 'DD-MM-YY'.")
        else:
            print("Inconsistent time format found in the following entries:")
            print(inconsistent_date_format['time'])

            for i, row in inconsistent_date_format.iterrows():
                if i >= max_rows_to_display:
                    break
                # Convert the row to a list and format it as needed for your error_log
                row_data = row.tolist()
                date_time_log.update([row_data], f'A{start_row}')  # Update error log with the row data


            print(f"Processing row {i+1} with data: {row_data}")

            # Move to the next row in the error log
            start_row += 1



        """

        # print(values_validated_df)
        print("\n\n\n\n\n")
        print(no_duplicates_df['time'])
        print("\n\n\n\n\n")

        """

        print("End of data validation\n")

        return no_duplicates_df


        

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
