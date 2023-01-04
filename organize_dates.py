"""
Code to organize Lisa's study dates
"""
from datetime import datetime
import pandas as pd

def import_data(): 
    # Specify the excel file name to be read in
    fname = "LEO Aim 3 - CRC Project Breeding & Test Dates.xlsx"
    
    # List of column names we will read in from the excel sheets
    study_date_columns = ['Mouse ID', 'Wean By', 'HFD/CON Start', 'HFD End', 'Pellet Coll Date 1', 'AOM D1', 'AOM D2', 'AOM D3', 'AOM D4', 'Pellet Coll Date 2', 'Ex Start Date (Calc)', 
                          'Pellet Coll Date 3', 'Sac Day', 'Age 5wk (#1)', 'Age 9wk (#2)', 'Age 13wk (#3)', 'Age 18wk (#4)', 'Age 21wk (#5 - EX wk1)', 'Age 25wk (#6 - EX wk5)', 
                          'Age 29wk (#7 - EX wk9)', 'Age 33wk (#8 - EX wk13)', 'Age 37wk (#9 - EX wk17)', 'Age 41wk (#10 - EX wk21)', 'Age 45wk (#11 - EX wk25)']

    mouse_weight_columns = ['Mouse ID','Age 5wk', 'Age 7wk', 'Age 9wk', 'Age 11wk', 'Age 13wk', 'Age 15wk', 'Age 18wk', 'Age 19wk', 'Age 21wk', 'Age 23wk', 'Age 25wk', 'Age 27wk', 
                            'Age 29wk', 'Age 31wk', 'Age 33wk', 'Age 35wk', 'Age 37wk', 'Age 39wk', 'Age 41wk', 'Age 43wk', 'Age 45wk', 'Age 47wk']
    
    # Read in the required sheets from the excel
    df1 = pd.read_excel(fname, sheet_name = 'Aim 1 Study Calendar - JV',  usecols = study_date_columns,   index_col = 'Mouse ID', engine='openpyxl')
    df2 = pd.read_excel(fname, sheet_name = 'Aim 2 Study Calendar - LEO', usecols = study_date_columns,   index_col = 'Mouse ID', engine='openpyxl')
    df3 = pd.read_excel(fname, sheet_name = 'Thicc Thursdays - JV',       usecols = mouse_weight_columns, index_col = 'Mouse ID', engine='openpyxl')
    df4 = pd.read_excel(fname, sheet_name = 'Thicc Thursdays - LEO',      usecols = mouse_weight_columns, index_col = 'Mouse ID', engine='openpyxl')    
    
    # Combine the sheets together and return single dataframe
    data = pd.concat([df1, df2, df3, df4], axis = 1)
    
    # Combine dublicate columns into one
    def sjoin(x): return ';'.join(x[x.notnull()].astype(str))
    data = data.groupby(level=0, axis=1).apply(lambda x: x.apply(sjoin, axis=1))
    
    # Convert columns from strings into datetime ignorie col 0 which is 'Mouse ID'
    data.iloc[:, :] = data.iloc[:, :].apply(pd.to_datetime, format='%Y-%m-%d', errors='coerce')

    return data

def date_window():
    
    # Prompt user for the start date of the interval
    date_1 = input("Enter start date in YYYY-MM-DD: ")
    
    # Prompt user for the end date of the interval
    date_2 = input("Enter end date in YYYY-MM-DD: ")
    
    # Convert the user provided strings into datatime data type
    date_1 = datetime.strptime(date_1, '%Y-%m-%d').date()
    date_2 = datetime.strptime(date_2, '%Y-%m-%d').date()
    
    # Return the two datetime objects
    return date_1, date_2

def get_dates(data, date_1, date_2):
    
    # Create a dictionary we shall fill with the dates and mouse IDs
    dates = {}
    
    # Loop for every column
    for column in data.columns:
        
        # Nested loop for every row of column
        for row in range(len(data)):

            # If the date in the excel is between the provided dates (inclusive interval)
            if (date_1 <= data[column].iloc[row] <= date_2) and (data[column].iloc[row] < data['Sac Day'].iloc[row]):
                
                # If that date is already in dictionary we will append a new mouse id and the task (column name)
                if data[column].iloc[row] in dates:
                    if column in dates[data[column].iloc[row]]:
                        dates[data[column].iloc[row]][column].append(str(int(data.index.values[row])))
                    else:
                        dates[data[column].iloc[row]][column] = []
                        dates[data[column].iloc[row]][column].append(str(int(data.index.values[row])))
                    
                # If date not already in dictionary create a new key and append mouse id and corresponding task
                else:
                    dates[data[column].iloc[row]] = {}
                    if column in dates[data[column].iloc[row]]:
                        dates[data[column].iloc[row]][column].append(str(int(data.index.values[row])))
                    else:
                        dates[data[column].iloc[row]][column] = []
                        dates[data[column].iloc[row]][column].append(str(int(data.index.values[row])))
    # Return the created dictionary            
    return dates

def save_dict(dates):
    
    # Open a notepad titled 'mouse_dates'
    with open("mice_dates.txt", 'w') as f:
        
        # For each key (date) in dictionary
        for key in sorted(dates.keys()):
            
            # Write the date to the notepad
            f.write('\n' + str(key) + '\n')
            
            # Nested loop for each value in the dictionary key
            for value in dates[key]:
                
                # For value in dictionary key write the value to the notepad
                f.write('\t' + value + ': ')
                for mouse in dates[key][value]:
                    f.write(mouse + ', ')
                
                # Once all mise have been printed for that event, start a new line
                f.write('\n')

def main():
    # Load and clean data
    data = import_data()
    
    # Get the date window of interest from user
    date_1, date_2 = date_window()
    
    # Organize all the dates/mice into a dictionary
    dates = get_dates(data, date_1, date_2)
    
    # Save dates to a notepad
    save_dict(dates)
    
main()