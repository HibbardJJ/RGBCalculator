# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import matplotlib.pyplot as plot
import os
os.chdir('C:\\Users\\Undergrunt\\PTImages')
from PIL import Image

im = Image.open('PT-18-06-05-11-29-41.jpg', 'r')
red, green, blue = im.split()
#width, height = im.size
pixel_values = list(im.getdata())
counts = im.histogram()

red_counts = red.histogram()
green_counts = green.histogram()
blue_counts = blue.histogram()
'''plot.plot(red_counts)
plot.plot(green_counts)
plot.plot(blue_counts)'''

from openpyxl import load_workbook
workbook = load_workbook('Purple Trad E Field Experiments')
workbook.create_sheet(date)
