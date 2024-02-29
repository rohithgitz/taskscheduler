# Task Scheduler Tasks Information and Database Insertion Script

## Overview
This Python script connects to the Task Scheduler on a specified machine, retrieves information about tasks from a specified folder, and inserts this information into a SQL Server database. The script also logs relevant information to a log file and prints updates to the console.

## Prerequisites
1. **Python Libraries:** Ensure the required Python libraries are installed using the following:
    ```bash
    pip install comtypes pyodbc tabulate
    ```

2. **Task Scheduler Configuration:** Update the IP address of the Task Scheduler machine and the folder name to fetch tasks from in the `get_task_scheduler_tasks` function.

3. **Database Configuration:** Update the database configuration parameters in the `db_config` dictionary under the `main` function.

## Usage
Run the script by executing the following command in the terminal:
```bash
python script_name.py
```

# Configuration Details
 Logging: The script logs information to a file named script_logs.log in the same directory.
 Output File: Standard output is redirected to a file named output_log_<current_date>.txt in the same directory.
Database Connection: The script connects to a SQL Server database using the specified configuration parameters (db_config).
Task Scheduler Connection: The script connects to the Task Scheduler on a specified machine (scheduler.Connect('172.31.165.131')).
# Functions
get_task_scheduler_tasks(folder_name='athens'):

Retrieves information about tasks in the specified Task Scheduler folder.
Returns a list of dictionaries containing task information.
insert_into_database(conn, task_data):

Inserts task information into the specified SQL Server database.
Checks for existing records before insertion.
main():

The main execution function.
Establishes a connection to the database, fetches task information, prints a table of tasks, and inserts data into the database.
Logging
The script logs information at different levels (INFO, ERROR) to the script_logs.log file. Additionally, standard output is redirected to the output_log_<current_date>.txt file.

# Important Notes
Ensure that the necessary permissions and configurations are set for connecting to the Task Scheduler and the SQL Server database.
Modify the script to fit specific requirements, such as adjusting database table names or handling errors differently.
Always handle sensitive information like database credentials securely.
Make sure the required Python libraries are installed before running the script.
Feel free to customize and adapt the script according to your specific needs.
