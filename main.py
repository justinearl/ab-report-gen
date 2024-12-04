import openpyxl
from openpyxl.chart import BarChart, Reference
import pandas as pd
from datetime import datetime, timedelta


data = [
        ["About", 79, 67, 78],
        ["Visit the Apollo", 60, 90, 50],
        ["Ticket and Events", 90, 50, 49],
        ["Membership", 98, 12, 33],
        ["Signature Seats", 100, 90, 70],
        ["Giving", 50, 20, 90],
        ["Education", 50, 30, 70],
        ["School Tours", 50, 70, 30],
        ["School Programs", 50, 80, 40],
        ["Apollo Theater Academy", 50, 30, 50],
        ["Rent the Apollo", 50, 100, 20],
        ["Tours", 50, 0, 30],
        ["Press Page", 50, 2, 90],
        ["Amateur Night", 50, 69, 70],
    ]

def generateExcel(data: list[list]):
    # Create a workbook and a worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Data"

    # Add data to the worksheet

    for row in data:
        ws.append(row)

    def create_chart(mode: str, column: int, position: str):
        # Create a bar chart
        chart = BarChart()
        chart.title = f"Desktop - Performance Report ({mode})"
        chart.x_axis.title = "Coverage"
        chart.y_axis.title = "Metrics"

        chart.legend = None

        chart.height = 10

        chart.y_axis.scaling.min = 0  
        chart.y_axis.scaling.max = 100

        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.y_axis.majorGridlines = None 

        # Add data to the chart
        maxRow = len(data)

        categories = Reference(ws, min_col=1, min_row=1, max_row=maxRow)
        values = Reference(ws, min_col=column, min_row=1, max_row=maxRow)
        chart.add_data(values)
        chart.set_categories(categories)

        # Place the chart on the worksheet
        ws.add_chart(chart, position)

    create_chart("Desktop", 2, "E5")
    create_chart("Mobile", 3, "N5")
    create_chart("Tablet", 4, "W5")

    wb.save("performance.xlsx")
    print("Excel file with graph created successfully!")

def getPerformanceDataInfile(file: str) -> dict[str, list[int]]:
    result = {}
    excelData = pd.read_excel(file)
    result["desktop"] = excelData["Unnamed: 1"].iloc[6:20].to_list()
    result["mobile"] = excelData["Unnamed: 5"].iloc[4:18].to_list()
    result["tablet"] = excelData["Unnamed: 9"].iloc[4:18].to_list()
    return result

def getDatesToExtract() -> list[str]:
    # Get the current date
    today = datetime.today()

    # Get the start of the current week (Monday)
    start_of_week = today - timedelta(days=today.weekday())

    # List to store weekdays in the required format
    weekdays = []

    # Loop through Monday to Friday (exclude weekends)
    for i in range(5):  # Only weekdays (0-4)
        day = start_of_week + timedelta(days=i)
        weekdays.append(day.strftime('%m-%d-%Y'))

    return weekdays

generateExcel(data)
