from datetime import datetime, timedelta, date, time
import argparse
from enum import Enum
import dateutil.parser
import xlsxwriter
import sqlite3
import json

class ReportType(Enum):
    """Constant definitions of report types"""
    now_dt = datetime.now()
    today_dt = datetime.combine(date.today(),time(0,0,0))
    H = {
        "name":"Values of Last 24 Hours",
        "abbrevation": "H",
        "type": "value",
        "static_dt":True,
        "start_dt": datetime.now()-timedelta(days=1),
        "end_dt":  now_dt
    }
    Y = {
        "name":"Values of Yesterday",
        "abbrevation": "Y",
        "type": "value",
        "static_dt":True,
        "start_dt": today_dt-timedelta(days=1),
        "end_dt":  today_dt
    }
    V = {
        "name":"Values of Period",
        "abbrevation": "V",
        "type": "value",
        "static_dt":False,
        "start_dt": today_dt-timedelta(days=7),
        "end_dt":  now_dt
    }
    P = {
        "name":"Analysis of Period",
        "abbrevation": "P",
        "type": "daily",
        "static_dt":False,
        "start_dt": today_dt-timedelta(days=7),
        "end_dt":  now_dt
    }
    S = {
        "name":"Analysis of 7 Days",
        "abbrevation": "S",
        "type": "daily",
        "static_dt":True,
        "start_dt": today_dt-timedelta(days=7),
        "end_dt":  now_dt
    }
    W = {
        "name":"Analysis of Last Week",
        "abbrevation": "W",
        "type": "daily",
        "static_dt":True,
        "start_dt": today_dt-timedelta(days=today_dt.weekday())-timedelta(days=7),
        "end_dt":  today_dt-timedelta(days=today_dt.weekday())
    }

class ReportSheet:
    """
    Report class for creating report sheets.
    """
    def __init__(self,name,abbrevation,type,start_dt,end_dt,db_file_name,table_name,table_type,ts_column_name,column_name,value_name):
        """
        initializating the class by report name, abbrevation, type, start and stop datetime.
        It also gets database file name, table name, table type, timestamp column name, value names and corresponding column name
        """
        self.name = name
        self.abbreviation = abbrevation
        self.type = type
        self.start_dt = start_dt
        self.end_dt = end_dt
        self.db_file_name = db_file_name
        self.table_name = table_name
        self.table_type = table_type
        self.ts_column_name = ts_column_name
        self.column_name = column_name
        self.value_name = value_name
        if self.type=="value":
            self.query = self.value_query()
        if self.type=="daily":
            self.query = self.daily_analysis()
    
    def __str__(self):
        return f'ReportSheet(\n\tname: {self.name}\n\tabbrevation: {self.abbreviation}\n\ttype: {self.type}\n\tstart_dt: {self.start_dt}\n\tend_dt: {self.end_dt})'

    def value_query(self):
        """Making query for values"""
        query = "SELECT strftime('%d-%m-%Y %H:%M',"+ self.ts_column_name +",'localtime') AS ts, "
        query += self.value_name + " "
        query += "FROM " + self.table_name + " "
        query += "WHERE datetime(" + self.ts_column_name + ",'localtime')<datetime('"+datetime.strftime(self.end_dt,"%Y-%m-%d %H:%M:%S")+"') "
        query += "AND "
        query += "datetime(" + self.ts_column_name + ",'localtime')>=datetime('"+datetime.strftime(self.start_dt,"%Y-%m-%d %H:%M:%S")+"')"
        return query


    # Check the query
    def daily_analysis(self):
        """Making query for daily analysis"""
        query = "SELECT "
        query += "q1.ts, "
        query += "q1.max_value, "
        query += "q1.avg_value, "
        query += "q1.min_value, "
        query += "q2.eod_value "
        query += "FROM "
        query += "(SELECT "
        query += "strftime('%d-%m-%Y'," + self.ts_column_name + ",'localtime') AS ts"
        query += "round(max(" + self.value_name + "),2) AS max_value, "
        query += "round(avg(" + self.value_name + "),2) AS avg_value, "
        query += "round(min(" + self.value_name + "),2) AS avg_value, "
        query += "max(" + self.ts_column_name + ") AS max_ts "
        query += "FROM " + self.table_name + " "
        query += "WHERE "
        query += "datetime(" + self.ts_column_name + ",'localtime')>=datetime('2022-08-01') "
        query += "AND "
        query += "datetime(" + self.ts_column_name + ",'localtime')<datetime('2022-08-07') "
        query += "GROUP BY ts) q1 INNER JOIN "
        query += "(SELECT "
        query += "datetime(" + self.ts_column_name + ",'localtime') AS ts, "
        query += "round(" + self.value_name + ",2) AS eod_value "
        query += "FROM " + self.table_name + ") q2 "
        query += "ON q1.max_ts=q2.ts"
        return query

def main():
    """Main program block"""

    # Getting starting datetime
    begin_dt = datetime.now()

    # Getting arguments
    arg_parser = argparse.ArgumentParser(
        prog="sqlitexlsx.py", 
        description="Creating excel report from arguments"
        )
    
    arg_parser.add_argument('-v','--version',action='version',version='%(prog)s 1.0')
    # args.mode must contain these characters: "H", "Y", "V", "P", "S" or 
    #  "W". If not, default is "H"
    arg_parser.add_argument('-m', '--mode', help="Report mode selection", default="H")
    arg_parser.add_argument('-s', '--start_dt', help="Starting datetime of period (e.g. 2021-12-31 or 2021-12-31 12:25:20")
    arg_parser.add_argument('-e', '--end_dt', help="End datetime of period (e.g. 2021-12-31 or 2021-12-31 12:25:20")
    arg_parser.add_argument('-d', '--data', help="Data name for data (e.g. rate or temperature...)")
    args = arg_parser.parse_args()

    # Getting configurations from settings.json
    settings_file = open('settings.json',"r")
    settings_text = settings_file.read()
    settings_file.close()
    settings_dict = json.loads(settings_text)
    del settings_file
    del settings_text

    # Checking if settings are usable
    # settings.json content must be like: 
    # {
    #   "rate": {
    #     "db_file_name": "data.db",
    #     "table_name": "usdTlRate",
    #     "table_type": "type1",
    #     "ts_column_name": "stamp",
    #     "column_name": ""
    #   }
    # }
    
    # argument for selecting variable is args.values for generating reports,
    # we are collecting them here
    variable_settings = settings_dict[args.data]
    del settings_dict

    # Reports to generate
    reports = []

    for key_char in ReportType:
        if key_char.name in args.mode:
            # Parsing the end_dt and start_dt. It uses static start_dt and 
            # end_dt in some report types, such as H or W. 
            # In period analysis report types:
            # 1- If there is start_dt and end_dt strings, it uses them. 
            # 2- If there is no start_dt string, default start_dt and end_dt 
            # values are used.
            # 3- If there is start_dt string in arguments, but no end_dt
            # string, it uses now_dt() datetime object in ReportTypes
            # enum
            if key_char.value["static_dt"]:
                report_start_dt = key_char.value["start_dt"]
                report_end_dt = key_char.value["end_dt"]
            else:
                if args.start_dt != None:
                    report_start_dt = dateutil.parser.parse(args.start_dt)
                    if args.end_dt != None:
                        report_end_dt = dateutil.parser.parse(args.end_dt)
                    else:
                        report_end_dt = ReportType.now_dt
                else:
                    report_start_dt = key_char.value["start_dt"]
                    report_end_dt = key_char.value["end_dt"]
            # After setting start_dt and end_dt variables, crate the 
            # ReportSheet object and append it to reports list
            reports.append(
                ReportSheet(
                    name=key_char.value["name"],
                    abbrevation=key_char.value["abbrevation"],
                    type=key_char.value["type"],
                    start_dt=report_start_dt,
                    end_dt=report_end_dt,
                    value_name= args.data,
                    db_file_name= variable_settings["db_file_name"],
                    table_name= variable_settings["table_name"],
                    table_type= variable_settings["table_type"],
                    ts_column_name= variable_settings["ts_column_name"],
                    column_name= variable_settings["column_name"]
                )
            )
    # End of creating reports list to generate

    # Create the workbook and display formats
    str_workbook_name = 'Report'+datetime.strftime(begin_dt, '%y%m%d%H%M')+'.xlsx'
    workbook = xlsxwriter.Workbook(str_workbook_name)
    bold_format = workbook.add_format({
        'bold': True
        })
    bold_centered_format = workbook.add_format({
        'bold': True,
        'center_across': True
        })
    float_format = workbook.add_format({
        'num_format': '0.00'
        })
    datetime_format = workbook.add_format({
        'num_format': 'dd/mm/yy hh:mm',
        'align': 'left'
        })
    date_format = workbook.add_format({
        'num_format': 'dd/mm/yy',
        'align': 'left'
        })

    # for every report in reports list:    
    for sheet in reports:
        # Gather the information from sqlite db
        conn = sqlite3.connect(sheet.db_file_name)
        cur = conn.cursor()
        cur.execute(sheet.query)
        fetch = cur.fetchall()
        conn.close()
        del cur
        del conn

        # Create worksheet for report without spaces
        worksheet = workbook.add_worksheet(sheet.name.replace(" ",""))

        # Writing values to the cells. There is 2 types of sheet template:
        # 1- Value: only shows the values with dt stamps. 1 column for 
        # data
        # 2- Analysis: analysing data and shows end of date, max and min
        # values with date timestamp. 3 columns for data
        
        # starting iteration from 1 since A1 is 0,0 for library
        row = 1

        if sheet.type=="value":
            worksheet.set_column('A:A',16)
            worksheet.set_column('B:B',14)
            worksheet.write(0,0,"Date & Time",bold_format)
            worksheet.write(0,1,sheet.value_name,bold_centered_format)
            for data in fetch:
                dt_value = datetime.strptime(data[0],'%d-%m-%Y %H:%M')
                worksheet.write_datetime(row,0,dt_value,datetime_format)
                worksheet.write(row,1,data[1],float_format)
                row += 1
        else:
            worksheet.set_column('A:A',16)
            worksheet.set_column('B:B',19)
            worksheet.set_column('C:C',19)
            worksheet.set_column('D:D',19)
            worksheet.set_column('E:E',19)
            worksheet.write(0,0,"Date",bold_format)
            worksheet.write(0,1,"Maximum",bold_centered_format)
            worksheet.write(0,2,"Average",bold_centered_format)
            worksheet.write(0,3,"Minimum",bold_centered_format)
            worksheet.write(0,3,"End of Day",bold_centered_format)
            for data in fetch:
                dt_value = datetime.strptime(data[0],'%d-%m-%Y')
                worksheet.write_datetime(row,0,dt_value,date_format)
                worksheet.write(row,1,data[1],float_format)
                worksheet.write(row,2,data[2],float_format)
                worksheet.write(row,3,data[3],float_format)
                worksheet.write(row,4,data[4],float_format)
                row += 1
        
        #create chart
        if sheet.type=="value":
            chart = workbook.add_chart({
                "type": "line"
            })
            chart.set_title({
                "name": sheet.name
            })
            chart.set_y_axis({
                "name": "TL",
                "num_format": "0.00"
            })
            chart.set_x_axis({
                "text_axis": True
            })
            chart.set_legend({
                "none": True,
                "position": "none"
            })
            chart.set_size({
                "width": 850,
                "height": 490
            })

            chart.add_series({
                "name": "USD/TL Rate",
                "categories": "="+sheet.name.replace(" ","")+"!$A$2:$A$"+str(row),
                "values": "="+sheet.name.replace(" ","")+"!$B$2:$B$"+str(row),
                "line": {"color": "red"},
                "marker": {"type": "none"}
            })
            worksheet.insert_chart("C1",chart)
        
        if sheet.type=="daily":
            chart = workbook.add_chart({
                "type": "line"
            })
            chart.set_title({
                "name": sheet.name
            })
            chart.set_y_axis({
                "name": "TL",
                "num_format": "0.00"
            })
            chart.set_x_axis({
                "date_axis": True,
                "num_format": "dd/mm/yyyy"
            })
            chart.set_legend({
                "none": True,
                "position": "none"
            })
            chart.set_size({
                "width": 850,
                "height": 490
            })

            chart.add_series({
                "name": "USD/TL Rate",
                "categories": "="+sheet.name.replace(" ","")+"!$A$2:$A$"+str(row),
                "values": "="+sheet.name.replace(" ","")+"!$B$2:$B$"+str(row),
                "line": {"color": "red"},
                "marker": {"type": "none"}
            })
            worksheet.insert_chart("F1",chart)

    workbook.close()
    print(str_workbook_name)
    end_dt = datetime.now()
    difference_dt = begin_dt - end_dt
    print(f'Program executed in {difference_dt.total_seconds()} seconds')

if __name__ == "__main__":
    main()