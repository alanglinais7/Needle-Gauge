#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jan  8 14:40:43 2021

@author: austinlanglinais
"""
import xlsxwriter

workbook = xlsxwriter.Workbook('chart_gauge.xlsx')
worksheet = workbook.add_worksheet()

chart_doughnut = workbook.add_chart({'type': 'doughnut'})
chart_pie = workbook.add_chart({'type': 'pie'})

# Add some data for the Doughnut and Pie charts. This is set up so the
# gauge goes from 0-100. It is initially set at 75%.
worksheet.write_column('H2', ['Donut', 12.5, 12.5, 12.5, 12.5, 100])
worksheet.write_column('I2', ['Pie', 75, 1, '=200-I4-I3'])

# Configure the doughnut chart as the background for the gauge.
chart_doughnut.add_series({
    'name': '=Sheet1!$H$2',
    'values': '=Sheet1!$H$3:$H$6',
    'points': [
        {'fill': {'color': 'red'}},
        {'fill': {'color': '#FF6600'}},
        {'fill': {'color': '#99CC00'}},
        {'fill': {'color': 'green'}},
        {'fill': {'none': True}}],
})

# Rotate chart so the gauge parts are above the horizontal.
chart_doughnut.set_rotation(270)

# Turn off the chart legend.
chart_doughnut.set_legend({'none': True})

# Turn off the chart fill and border.
chart_doughnut.set_chartarea({
    'border': {'none': True},
    'fill': {'none': True},
})

# Configure the pie chart as the needle for the gauge.
chart_pie.add_series({
    'name': '=Sheet1!$I$2',
    'values': '=Sheet1!$I$3:$I$6',
    'points': [
        {'fill': {'none': True}},
        {'fill': {'color': 'black'}},
        {'fill': {'none': True}}],
})

# Rotate the pie chart/needle to align with the doughnut/gauge.
chart_pie.set_rotation(270)

# Combine the pie and doughnut charts.
chart_doughnut.combine(chart_pie)

# Insert the chart into the worksheet.
worksheet.insert_chart('A1', chart_doughnut)

workbook.close()