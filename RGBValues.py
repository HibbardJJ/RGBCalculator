# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import matplotlib.pyplot as plot
import os
os.chdir('C:\\Users\\Joshua Hibbard\\Box\\Josh data\\PT RGB\\04-20-18')
from PIL import Image
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
)
#experiment_date = input('What is the experiment date? Write in MM-DD-YY format.')

im = Image.open('NoE.jpg', 'r')
red, green, blue = im.split()
#width, height = im.size
#pixel_values = list(im.getdata())
#counts = im.histogram()

red_counts = red.histogram()
green_counts = green.histogram()
blue_counts = blue.histogram()
plot.figure(1)
plot.plot(red_counts)
plot.figure(2)
plot.plot(green_counts)
plot.figure(3)
plot.plot(blue_counts)
plot.show()
from openpyxl import load_workbook
workbook = load_workbook('Test.xlsx')
#image_list = os.listdir(str(experiment_date))
ws = workbook.create_sheet('04-18-20')
#Write the x-values
for i in range(256):
    ws.cell(row = i + 1, column = 1).value = i
    ws.cell(row = i + 1, column = 2).value = red_counts[i]
    ws.cell(row = i + 1, column = 3).value = green_counts[i]
    ws.cell(row = i + 1, column = 4).value = blue_counts[i]

chart = ScatterChart()
chart.title = "Red"
chart.style = 13
chart.x_axis.title = 'Pixel Bin'
chart.y_axis.title = 'Intensity'

xvalues = Reference(ws, min_col=1, min_row=1, max_row=256)
for i in range(2, 5):
    values = Reference(ws, min_col=i, min_row=1, max_row=256)
    series = Series(values, xvalues, title_from_data=True)
    chart.series.append(series)

ws.add_chart(chart, "A10")


workbook.save(filename='Test.xlsx')