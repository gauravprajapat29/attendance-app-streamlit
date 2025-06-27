import pandas as pd
import calendar
import warnings
import datetime
from datetime import datetime as dt
import numpy as np

# Ignore warning messages for clean output
warnings.filterwarnings("ignore")

def process_excel(file, year, month):
    # Load Excel sheet with attendance logs
    df = pd.read_excel(file, sheet_name="Att.log report", header=3)

    # Handle case where rows are odd (pair of IN/OUT for each employee)
    if df.shape[0] % 2 != 0:
        df.loc[len(df)] = [np.nan] * len(df.columns)

    # Extract employee names from every second row (index 0, 2, 4...)
    employee_names = [df.iloc[i, 10] for i in range(0, df.shape[0], 2)]

    # Drop IN rows; keep OUT rows only
    df.drop(index=[i for i in range(0, df.shape[0], 2)], inplace=True)

    # Add employee name column
    df['Emp Name'] = employee_names
    df.index = employee_names

    # Convert attendance string "09:10-18:30" to "09:10,18:30"
    def format_attendance(entry):
        if isinstance(entry, str):
            return entry[:5] + "," + entry[-5:]
        return entry

    for col in df.columns:
        if col != "Emp Name":
            df[col] = df[col].apply(format_attendance)

    # Remove rows with almost all missing values
    df = df[df.notnull().sum(axis=1) > 1]

    # Get Sundays in the given month
    def get_sundays(year, month):
        _, total_days = calendar.monthrange(year, month)
        return [day for day in range(1, total_days + 1)
                if datetime.date(year, month, day).weekday() == 6]

    sundays = get_sundays(year, month)

    # Initialize summary columns
    df['Sunday'] = len(sundays)
    df['Full Day'] = 0
    df['Late In'] = 0
    df['Half Day'] = 0
    df['Early Out'] = 0
    df['Leave'] = 0
    df['Missed In Out'] = 0

    # Attendance logic for trainers
    def calculate_trainer_metrics(entries):
        in_time_std = dt.strptime("09:15", "%H:%M")
        out_time_std = dt.strptime("17:31", "%H:%M")
        late_in_threshold = dt.strptime("09:45", "%H:%M")
        full_out_threshold = dt.strptime("18:20", "%H:%M")
        half_day_start = dt.strptime("14:00", "%H:%M")
        half_day_end = dt.strptime("16:00", "%H:%M")
        early_out_start = dt.strptime("16:00", "%H:%M")
        early_out_end = dt.strptime("17:30", "%H:%M")

        half_day = late = full_day = early_out = leave = missed = 0

        for entry in entries:
            if isinstance(entry, str):
                in_str, out_str = entry.split(",")
                in_time = dt.strptime(in_str, "%H:%M")
                out_time = dt.strptime(out_str, "%H:%M")
                work_hours = (out_time - in_time).seconds // 3600

                if work_hours >= 3:
                    if half_day_start <= out_time <= half_day_end:
                        half_day += 1
                    if in_time > in_time_std and (out_time_std <= out_time < full_out_threshold):
                        late += 1
                        full_day += 1
                    elif in_time > late_in_threshold:
                        late += 1
                    if early_out_start <= out_time <= early_out_end:
                        early_out += 1
                        full_day += 1
                    if in_time <= in_time_std and out_time >= out_time_std:
                        full_day += 1
                    elif in_time <= late_in_threshold and out_time >= full_out_threshold:
                        full_day += 1
                else:
                    missed += 1
            else:
                leave += 1

        return half_day, late, full_day, early_out, leave, missed

    # Attendance logic for non-trainers
    def calculate_non_trainer_metrics(entries):
        in_time_std = dt.strptime("09:45", "%H:%M")
        out_time_std = dt.strptime("18:21", "%H:%M")
        half_day_start = dt.strptime("14:00", "%H:%M")
        half_day_end = dt.strptime("16:00", "%H:%M")
        early_out_start = dt.strptime("16:00", "%H:%M")
        early_out_end = dt.strptime("18:20", "%H:%M")

        half_day = late = full_day = early_out = leave = missed = 0

        for entry in entries:
            if isinstance(entry, str):
                in_str, out_str = entry.split(",")
                in_time = dt.strptime(in_str, "%H:%M")
                out_time = dt.strptime(out_str, "%H:%M")
                work_hours = (out_time - in_time).seconds // 3600

                if work_hours >= 3:
                    if half_day_start <= out_time <= half_day_end:
                        half_day += 1
                    if in_time > in_time_std:
                        late += 1
                    if early_out_start <= out_time <= early_out_end:
                        early_out += 1
                        full_day += 1
                    if out_time >= out_time_std:
                        full_day += 1
                else:
                    missed += 1
            else:
                leave += 1

        return half_day, late, full_day, early_out, leave, missed

    # Define trainer names (can be replaced by a config or database)
    trainer_names = ["Ritesh Naidu", "sagar", "Sidharth", "ZuLfikar", "GauravPrajapat"]

    # Loop through each employee and calculate metrics
    for i, name in enumerate(df.index):
        employee_data = df.iloc[i, :-8]

        if name in trainer_names:
            metrics = calculate_trainer_metrics(employee_data)
        else:
            metrics = calculate_non_trainer_metrics(employee_data)

        df.at[name, "Half Day"] = metrics[0]
        df.at[name, "Late In"] = metrics[1]
        df.at[name, "Full Day"] = metrics[2]
        df.at[name, "Early Out"] = metrics[3]
        df.at[name, "Leave"] = metrics[4]
        df.at[name, "Missed In Out"] = metrics[5]

    # Subtract Sundays from total leaves
    df['Leave'] = df['Leave'] - df['Sunday']

    # Return final summary
    return df.iloc[:, -7:].reset_index()
