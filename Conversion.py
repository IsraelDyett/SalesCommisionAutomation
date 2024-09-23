import pyodbc
import datetime
import os
import logging
import argparse

# Set up logging
#log_directory = r'C:\DevelopmentApps\Commissions Comparison Text File\Logs'
log_directory = r'C:\Temp\Commissions\CommissionsLogs'
os.makedirs(log_directory, exist_ok=True)
log_filename = os.path.join(log_directory, 'error_log.txt')

logging.basicConfig(filename=log_filename,
                    level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Set up argument parsing
parser = argparse.ArgumentParser(description='Execute SQL query and save results to a tab-delimited text file.')
parser.add_argument('--overWriteMonthAndYear', type=int, default=0, help='Value for @overWriteMonthAndYear (default: 0)')
parser.add_argument('--FY', type=str, default=None, help='Value for @FY (default: NULL)')
parser.add_argument('--CalMnth', type=str, default=None, help='Value for @CalMnth (default: NULL)')
parser.add_argument('--directory', type=str, default=r'\\10.120.110.52\Payroll Docs\Payroll Shared\HRD-Salesrep Commission Calculation\Downloaded Data', help=r'Value for directory (default: \\10.120.110.52\Payroll Docs\Payroll Shared\HRD-Salesrep Commission Calculation\Downloaded Data)')
#parser.add_argument('--directory', type=str, default=r'\\10.120.110.52\Payroll Docs\Payroll Shared\HRD-Salesrep Commission Calculation\testing - Gabby', help=r'Value for directory (default: \\10.120.110.52\Payroll Docs\Payroll Shared\HRD-Salesrep Commission Calculation\testing - Gabby)')
#parser.add_argument('--directory', type=str, default=r'C:\DevelopmentApps', help=r'Value for directory (default: C:\DevelopmentApps)')

args = parser.parse_args()

try:
	# Log the provided arguments
	logging.error(f"""
	@overWriteMonthAndYear = {args.overWriteMonthAndYear},
	@FY = {args.FY},
	@CalMnth = {args.CalMnth},
	Directory = {args.directory}
	""")
    
	# Define the connection parameters
	server = '10.120.110.44'
	database = 'MDData'
	username = 'sa'
	password = 'install#07'
	driver = '{ODBC Driver 17 for SQL Server}'  # Ensure the correct ODBC driver is installed

	# Establish the database connection
	conn = pyodbc.connect(
		f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

	# Create a cursor object
	cursor = conn.cursor()

	# Define the SQL query
	sql_query = f"""EXEC	 [dbo].[SalesCommissionsTxt]
			@overWriteMonthAndYear = {args.overWriteMonthAndYear},
			@FY = {'NULL' if args.FY is None else args.FY},
			@CalMnth = {'NULL' if args.CalMnth is None else args.CalMnth}
	"""

	# Execute the SQL query
	cursor.execute(sql_query)

	# Fetch all rows from the executed query
	rows = cursor.fetchall()
	fy = args.FY 
	cal_mnth = args.CalMnth
	directory = args.directory
      
	# Determine the output file name based on parameters or default to current date
	if fy and cal_mnth:
		year = fy
		formatted_month = f'{int(cal_mnth):02d}'
	else:
		now = datetime.datetime.now()
		previous_month = now.month - 1 if now.month > 1 else 12
		year = now.year if now.month > 1 else now.year - 1
		formatted_month = f'{previous_month:02d}'

	# Create the output file name
	output_filename = f'Slsqts{year}_{formatted_month}.txt'

	#directory = r'C:\DevelopmentApps'
	#directory = r'\\10.120.110.52\Payroll Docs\Payroll Shared\HRD-Salesrep Commission Calculation\testing - Gabby'
	#directory = r'\\10.120.110.52\Payroll Docs\Payroll Shared\HRD-Salesrep Commission Calculation\Downloaded Data'

	# Ensure the directory exists
	os.makedirs(directory, exist_ok=True)

	# Combine directory path with filename
	full_path = os.path.join(directory, output_filename)

	# Write the data to a text file in tab-delimited format
	with open(full_path, 'w') as f:
		for row in rows:
			f.write('\t'.join(str(field) for field in row) + '\n')

except pyodbc.Error as e:
    logging.error(f"Database error: {e}")
except Exception as e:
    logging.error(f"An error occurred: {e}")
finally:
    # Close the cursor and connection
    try:
        cursor.close()
        conn.close()
    except Exception as e:
        logging.error(f"Error closing connection: {e}")