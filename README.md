# Creating Excel Reports
## Introduction
In this report manufacturing python program, I am aiming to collecting data from sqlite database then parsing to the excel sheet by the report type.

## Needed libraries
Needed libraries to execute program are.
- datetime (standard library)
- argparse (standard library)
- enum (standard library)
- dateutil (needed to install)
- xlsxwriter (needed to install)
- sqlite3 (standard library)
- json (standard library)

To install these libraries use the requirements.txt file at the project folder as:
```
pip install -r requirements.txt
```

## Main Usage
Main usage is done by terminal command below:
```
python3 makexlsx.py -d [data _name] -m [mode] -s [start datetime] -e [end datetime]
```
If we look for arguments:

### 1- `-d [data name]` Argument
Data name argument is for selecting and calculating from one of multiple data variables. This argument is the only mandatory argument. 

Data variables' accessing attributes are defined at `settings.json` file. A `settings.json` file is defined and can be usable for test purposes. For test purposes currently usable values are:
- rate
- humidity
- temperature
- pressure

### 2- `-m [mode]` Argument
Mode argument is for report types. There is six report types currently. It is defined by enum class in the beginning of program. Abbrevation of report types must be written as a string in this argument. Order of the abbrevation doesn't affect the report sheets.

Example usage:
```
python3 makexlsx.py -m HYVPSW -d [data_name]
```
#### a. `H` Argument: Values of Last 24 Hours
In `H` argument, program gathers values from last 24 hours and parses to the sheet. 

In `H` argument, program doesn't looks for `start_dt` or `end_dt`

#### b. `Y` Argument: Values of Yesterday
In `Y` argument, program gathers values from yesterday and parses to the sheet. 

In `Y` argument, program doesn't looks for `start_dt` or `end_dt`

#### c. `V` Argument: Values of Period
In `V` argument, program gathers values between `start_dt` and `end_dt` and parses to the sheet. 

In `V` argument, program looks for `start_dt` and `end_dt` variables. If there is no arguments for `start_dt` and/or `end_dt`, program uses current date and time for `end_dt` and 7 days ago for `start_dt`.

#### d. `P` Argument: Daily Analysis of Period
In `P` argument, program gathers values between `start_dt` and `end_dt`. After analyses for maximum value, minimum value, average value and end of day value, program parses the daily analysis values to sheet.

In `V` argument, program looks for `start_dt` and `end_dt` variables. If there is no arguments for `start_dt` and/or `end_dt`, program uses current date and time for `end_dt` and 7 days ago for `start_dt`.

#### e. `S` Argument: Daily Analysis of Last 7 Days
In `S` argument, program gathers values between current date time and start of day of 7 days ago. After analyses for maximum value, minimum value, average value and end of day value, program parses the daily analysis values to sheet. 

In `S` argument, program doesn't looks for `start_dt` and `end_dt` variables. Automatically gathers the `start_dt` by calculating start of day of 7 days ago and `end_dt` by current date time.

#### f. `W` Argument: Daily Analysis of Last Week
In `W` argument, program gathers values between Monday and Sunday of last week. After analyses for maximum value, minimum value, average value and end of day value, program parses the daily analysis values to sheet. 

In `W` argument, program doesn't looks for `start_dt` and `end_dt` variables. Automatically gathers the `start_dt` by calculating start of day of last weeks Monday and `end_dt` by end of day of last weeks Sunday.

### 3- `-s [start datetime]` Argument
In `-s` argument, argument string for start datetime is typed such as:

```
-s 2022-08-01 12:15:00
```
or
```
-s 2022-08-01
```

By the `dateutil` library, program gathers both date and datetime strings and converts to datetime objects.

### 4- `-e [end datetime]` Argument
After `-e` argument, argument string for start datetime is typed such as:

```
-e 2022-08-01 12:15:00
```
or
```
-e 2022-08-01
```

By the `dateutil` library, program gathers both date and datetime strings and converts to datetime objects.
## Creating Excel Files
Program creates `.xlsx` excel file as report by the arguments given. To explain in steps:
1. By data argument, program reaches the related database and table.
2. By report mode and/or end and start datetimes, program creates report sheet object list that contains the necessary information to gather values or analysis from database and sheet template
3. After creating report sheet object list, program creates a `.xlsx` file. File name contains the date and time program started. Such as: 
```
Report20220801101234.xlsx
Report+[year]+[month]+[day]+[hour]+[minutes]+[seconds]+.xlsx
```
4. Program starts a loop for creating sheets for every report mode for corresponding data variable by the report sheet list. In this list, program gathers values from sqlite database and parses to excel sheet then draws a chart by the parsed values in excel sheet.
5. After creating sheets, program closes excel file and prints the excel file name to the terminal to use file later.