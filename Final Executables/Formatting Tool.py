#Developed by Johann Vivas on 03/24/2023

from pandas import read_excel, DataFrame, to_datetime
import datetime as dt
import pathlib
from warnings import filterwarnings
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def replaceDates(new_df):
    new_date_list = []
    for d in new_df['Start Date']:
        #print(dt.datetime.strptime(d, r"%d/%m/%Y").date())
        new_d_obj = dt.datetime.strptime(d, r"%Y-%m-%d").date()
        month = new_d_obj.month
        day = new_d_obj.day
        year = new_d_obj.year
        new_d = f"{month}/{day}/{year}"
        new_date_list.append(new_d)

    new_df['Start Date'] = new_date_list

def replaceLunchDur(new_df):
    for i in range(0, len(new_df["Description"])):
        if new_df["Description"].loc[i] == "Lunch": new_df['Duration (h)'].loc[i] = ""

filterwarnings('ignore', category=UserWarning, module="openpyxl")

print("Please select your downloaded (unedited) intern productivity .xlsx file.")
Tk().withdraw()
filename = askopenfilename()

if pathlib.Path(filename).suffix != ".xlsx":
    input("Invalid file format. Pass only .xlsx files to this application. Press enter to exit.")
    raise Exception("Invalid File Format")
    
try:
    dataframe = read_excel(filename)
    new_df = dataframe.copy()

    new_df = new_df.drop(['Client', 'Task', 'User', 'Group', 'Email',
                        'Billable', 'End Date', 'Duration (decimal)'], axis=1, inplace=False)

    column_order = ['Start Date', 'Project', 'Tags', 'Start Time', 'End Time', 'Duration (h)', 'Description']
    new_df = new_df.reindex(columns=column_order)

    new_df['Start Date'] = to_datetime(new_df['Start Date'], format=r"%d/%m/%Y")
    new_df = new_df.sort_values(by=['Start Date', 'Start Time'],ascending=True)
    new_df['Start Date'] = new_df['Start Date'].astype(str)

    replaceDates(new_df=new_df)

    replaceLunchDur(new_df=new_df)

    final_df = DataFrame({
        'Date':new_df['Start Date'],
        'Project':new_df['Project'],
        'Tag':new_df['Tags'],
        'Start Time':new_df['Start Time'],
        'End Time':new_df['End Time'],
        'Duration':new_df['Duration (h)'],
        'Description':new_df['Description']
    })

    finalfname = fr'formattedtimesheet({dt.datetime.today().strftime(r"%m-%d-%Y")}).xlsx'
    filepath = fr'./{finalfname}'

    final_df.to_excel(filepath, index=False, header=True)
    print(f"Formatting Done! Filename: {finalfname}")
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    input("Press enter to exit...")