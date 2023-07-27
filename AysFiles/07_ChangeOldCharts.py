# For this code we need an excel file. This excel file will include our chart datas. This code now draw just one line.

import openpyxl

# Load the new Excel file
wb = openpyxl.load_workbook('ornek.xlsx')

# Access the sheet containing the new data
sheet = wb['Tabelle1']

# Read the new categories and values from the Excel file
new_categories = [cell.value for cell in sheet['A'][1:]]
new_values = [cell.value for cell in sheet['B'][1:]]

from pptx import Presentation
from pptx.chart.data import ChartData

# Load the PowerPoint file
presentation = Presentation('ornek.pptx')

# Loop through each slide
for slide in presentation.slides:
    for shape in slide.shapes:
        # Check if the shape is a chart
        if shape.has_chart:
            chart = shape.chart

            # Create ChartData object with new categories and values
            chart_data = ChartData()
            chart_data.categories = new_categories
            chart_data.add_series('Updated Data', new_values)

            # Replace the chart data with the new data
            chart.replace_data(chart_data)

presentation.save('07_ChangeOldCharts.pptx')
