import os
import glob
import pathlib
import xlsxwriter
import pandas as pd
from datetime import datetime

parentPath = str(pathlib.Path(__file__).parent.absolute())

sourcePath = parentPath + '\\xlsx'
outputPath = parentPath + '\\xlsx'

os.makedirs(sourcePath, exist_ok=True)
os.makedirs(outputPath, exist_ok=True)

for filePath in glob.glob(sourcePath + '\\*.xlsx'):
    fileName = os.path.basename(filePath)[:-5]

    if fileName.endswith('--processed'):
        continue

    xls = pd.ExcelFile(sourcePath + '\\' + fileName + '.xlsx')
    wb = xlsxwriter.Workbook(outputPath + '\\' +
                             fileName + '--processed.xlsx')
    sheetIndex = 0
    for sheet in xls.sheet_names:
        if sheet == 'Index':
            continue

        df = pd.read_excel(sourcePath + '\\' +
                           fileName + '.xlsx', sheet_name=sheet)

        timeList = df[df.columns[0]].tolist()
        dateList = df['Date'].tolist()
        idList = df['ID'].tolist()
        latList = df['Latitude'].tolist()
        longList = df['Longitude'].tolist()

        ws = wb.add_worksheet(sheet)

        for row in range(len(timeList)):
            if type(timeList[row]) is str:
                timeList[row] = datetime.strptime(timeList[row], '%H:%M:%S')

        ws.set_column(0, 8, 10)
        ws.set_column(1, 1, 11)
        ws.set_column(3, 4, 11)
        ws.set_column(2, 2, 9)

        titleFormat = wb.add_format()
        titleFormat.set_bold()
        numberFormat = wb.add_format()
        numberFormat.set_num_format('0.000')
        numberFormatOne = wb.add_format()
        numberFormatOne.set_num_format('0.0')
        timeFormat = wb.add_format()
        timeFormat.set_num_format('hh:mm:ss')

        # ws.write(0, 0, 'Time', titleFormat)
        # ws.write(0, 1, 'Date', titleFormat)
        # ws.write(0, 2, 'ID', titleFormat)
        # ws.write(0, 3, 'Latitude', titleFormat)
        # ws.write(0, 4, 'Longitude', titleFormat)
        # ws.write(0, 5, 'dt (s)', titleFormat)
        # ws.write(0, 6, 'dx (m)', titleFormat)
        # ws.write(0, 7, 'Speed (m/s)', titleFormat)
        # ws.write(0, 8, 'Speed (km/h)', titleFormat)
        # ws.write(0, 9, 'Duration (hh:mm:ss)', titleFormat)
        # ws.write(0, 10, 'Distance (km)', titleFormat)
        # ws.write(0, 11, 'Avg. Speed (km/h)', titleFormat)

        title_list = ['Time', 'Date', 'ID', 'Latitude', 'Longitude', 'dt (s)', 'dx (m)', 'Speed (m/s)', 'Speed (km/h)', 'Duration (hh:mm:ss)', 'Distance (km)', 'Avg. Speed (km/h)']

        for index, title in enumerate(title_list):
            ws.write(0, index, title, titleFormat)

        for row in range(len(timeList)):
            ws.write(row + 1, 0, timeList[row], timeFormat)
            ws.write(row + 1, 1, dateList[row])
            ws.write(row + 1, 2, idList[row])
            ws.write(row + 1, 3, latList[row])
            ws.write(row + 1, 4, longList[row])

        for row in range(2, len(timeList) + 1):
            ws.write_formula(
                f'F{row + 1}', f'=IF(OR((A{row + 1}-A{row})*24*60*60 >= 30 , (A{row + 2}-A{row + 1})*24*60*60 >= 30), 0, (A{row + 1}-A{row})*24*60*60)')
            ws.write_formula(
                f'G{row + 1}', f'=IF(OR(F{row + 1} <= 0, D{row + 1} = 0, E{row + 1} = 0, D{row} = 0, E{row} = 0), 0, IF(AND(D{row + 1}=D{row},E{row + 1}=E{row}),0,ACOS(COS(RADIANS(90-D{row + 1}))*COS(RADIANS(90-D{row}))+SIN(RADIANS(90-D{row + 1}))*SIN(RADIANS(90-D{row}))*COS(RADIANS(E{row + 1}-E{row})))*6371*1000))', cell_format=numberFormat)
            ws.write_formula(
                f'H{row + 1}', f'=IF(F{row + 1} <= 0, 0, G{row + 1}/F{row + 1})', cell_format=numberFormat)
            ws.write_formula(
                f'I{row + 1}', f'=H{row + 1}*3.6', cell_format=numberFormat)

        ws.write_formula('J2', f'=SUM(F:F)/24/60/60',
                         cell_format=timeFormat)
        ws.write_formula('K2', f'=SUM(G:G)/1000',
                         cell_format=numberFormatOne)
        ws.write_formula('L2', f'=K2/J2/24',
                         cell_format=numberFormatOne)

        chart = wb.add_chart({'type': 'scatter',
                              'subtype': 'straight_with_markers'})

        chart.add_series({
            'name':       'Speed (km/h)',
            'categories': f'=\'{sheet}\'!$A$3:$A${len(timeList) + 1}',
            'values':     f'=\'{sheet}\'!$I$3:$I${len(timeList) + 1}',
        })

        chartMinX = timeList[0].hour
        chartMaxX = timeList[-1].hour + 1

        if timeList[0].minute <= 5:
            chartMinX -= 1
        if timeList[-1].minute >= 55:
            chartMaxX += 1

        chart.set_title(
            {'name': f'Speed vs. Time of Travel Log {sheet}'})
        chart.set_y_axis({'name': 'Speed (km/h)', 'max': 70, 'num_format': '0',
                          'major_gridlines': {'visible': True}})
        chart.set_x_axis({'name': 'Time (hh:mm)', 'num_format': 'hh:mm', 'major_unit': 1/24, 'minor_unit': 1/48, 'min': (chartMinX % 24)/24, 'max': (chartMaxX % 24)/24,
                          'major_gridlines': {'visible': True}})
        chart.set_legend({'none': True})
        chart.set_size({'width': 704, 'height': 400})

        ws.insert_chart('N2', chart)

        sheetIndex += 1
        print(
            f'{fileName} is being processed. ({int(sheetIndex*100/len(xls.sheet_names))}%)')

    try:
        wb.close()
        print(f'{fileName} was successfully processed.')
    except:
        print('Something went wrong while creating the processed XLSX file.')
