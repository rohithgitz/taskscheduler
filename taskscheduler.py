import sys
import comtypes.client
from datetime import datetime, timedelta
import pyodbc
from tabulate import tabulate
import logging

# Configure logging
logging.basicConfig(filename='script_logs.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Output file for additional information
output_file_path = f'output_log_{datetime.now().strftime("%Y-%m-%d")}.txt'

# Redirect standard output to a file
sys.stdout = open(output_file_path, 'w')

def get_task_scheduler_tasks(folder_name='athens'):
    tasks_info = []

    logging.info("Connecting to Task Scheduler...")
    try:
        scheduler = comtypes.client.CreateObject('Schedule.Service')
        scheduler.Connect('172.31.165.131')  # Change to your IP address
        logging.info("Connected to Task Scheduler.")
    except comtypes.COMError as e:
        logging.error(f"Error connecting to Task Scheduler: {e}")
        print(f"{datetime.now()} - Error connecting to Task Scheduler: {e}")
        return tasks_info

    root_folder = scheduler.GetFolder('\\')
    
    try:
        selected_folder = root_folder.GetFolder(folder_name)
    except comtypes.COMError as e:
        logging.error(f"Error accessing folder '{folder_name}' in Task Scheduler: {e}")
        print(f"{datetime.now()} - Error accessing folder '{folder_name}' in Task Scheduler: {e}")
        return tasks_info

    logging.info(f"Fetching tasks from Task Scheduler folder '{folder_name}'...")
    tasks = selected_folder.GetTasks(0)  # 0 represents all tasks

    for task in tasks:
        task_info = {
             'task_name': task.Name,
            'task_path': task.Path,
            'status': task.State,
            'last_run_time': datetime(1899, 12, 30) + timedelta(days=task.LastRunTime),
            'last_task_result': task.LastTaskResult,
            'next_run_time': task.NextRunTime,
            'author': task.Definition.Principal.UserId,        }
        tasks_info.append(task_info)

    logging.info(f"Tasks from folder '{folder_name}' fetched successfully.")
    #for task in tasks:
    #logging.info(f"Task Name: {task.Name}, Last Run Time: {task.LastRunTime}")

    return tasks_info

def insert_into_database(conn, task_data):
    cursor = conn.cursor()

    for task in task_data:
        try:
            # Check if the task_name already exists in the database
          
            cursor.execute('SELECT COUNT(*) FROM completed_tasks WHERE task_name = ? AND CAST(last_run_time AS DATE) =?', (task['task_name'],task['last_run_time'].strftime('%Y-%m-%d')))
            #print ('task_name', 'last_run_time')
            #logging.info(f"Task Name: {task['task_name']}, Last Run Time: {task['last_run_time']}")
            logging.info(f"Task Name: {task['task_name']}, Last Run Time: {task['last_run_time'].strftime('%Y-%m-%d')}")

            existing_records = cursor.fetchone()[0]

            if existing_records == 0:
                # Task does not exist in the database, proceed with insertion
                cursor.execute('''
                   INSERT INTO completed_tasks (
                        task_name, task_path, status, last_run_time, last_task_result, next_run_time, author
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    task['task_name'],
                    task['task_path'],
                    task['status'],
                    task['last_run_time'],
                    0 if task['last_task_result'] == 0 else 1,
                    task['next_run_time'],
                    task['author']
                ))
                conn.commit()
                logging.info(f"Data inserted into the database for task: {task['task_name']}")
                print(f"{datetime.now()} - Data inserted for task: {task['task_name']}")
            else:
                logging.info(f"Task '{task['task_name']}' already exists in the database. Skipping insertion.")
                print(f"{datetime.now()} - Task '{task['task_name']}' already exists in the database. Skipping insertion.")
        except pyodbc.Error as e:
            logging.error(f"Error inserting data for task {task['task_name']}: {e}")
            print(f"{datetime.now()} - Error inserting data for task {task['task_name']}: {e}")


def main():
    logging.info("Script execution started.")
    print(f"{datetime.now()} - Script execution started.")

    # Connect to SQL Server database
    db_config = {
        'server': '10.1.12.14',
        'database': 'PowerBI_Stage',
        'user': 'pp_sveerla',
        'password': '!Welcome@2023SV$$',
        'driver': '{ODBC Driver 17 for SQL Server}',  # Adjust the driver based on your SQL Server version
    }

    try:
        conn = pyodbc.connect(**db_config)
        logging.info("Connected to the database.")
        print(f"{datetime.now()} - Connected to the database.")
      #  truncate_completed_tasks_table(conn)  # Truncate the table
    except pyodbc.Error as err:
        logging.error(f"Error connecting to the database: {err}")
        print(f"{datetime.now()} - Error connecting to the database: {err}")
        return

    # Fetch task names dynamically from the Task Scheduler
    tasks_info = get_task_scheduler_tasks(folder_name='athens')

    if tasks_info:
        table_data = []
        for task in tasks_info:
            table_data.append([
                task['task_name'],
                task['task_path'],
                task['status'],
                task['last_run_time'],
                task['last_task_result'],
                task['next_run_time'],
                task['author']
            ])

        headers = ["Task Name", "Task Path", "Status", "Last Run Time", "Last Task Result", "Next Run Time", "Author"]
        table = tabulate(table_data, headers=headers, tablefmt="pretty")
        print(table)

        insert_into_database(conn, tasks_info)
        logging.info("Data insertion process completed.")
        print(f"{datetime.now()} - Data insertion process completed.")
    else:
        logging.info("No tasks found in Task Scheduler folder 'athens'.")
        print(f"{datetime.now()} - No tasks found in Task Scheduler folder 'athens'.")

    # Close the database connection when done
    conn.close()
    logging.info("Script execution completed.")
    print(f"{datetime.now()} - Script execution completed.")

if __name__ == "__main__":
    main()
