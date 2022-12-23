import os
import glob
import pathlib
import time
import json
import xlsxwriter
import pandas as pd
from datetime import datetime
from termcolor import colored

def close_workbook(wb, file_name, start_time):
    print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
          colored('Saving the XLSX file...', 'grey'))
    try:
        wb.close()
        print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
              colored(f'{file_name}.xlsx',
                      'green'), 'was successfully exported in',
              colored(f'{time.time() - start_time:.2f} seconds.', 'red'))
    except:
        print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'),
              'cyan'), 'Something went wrong while saving the XLSX file.')
    print('-' * 150)


database_key = 'Bus'
single_workbook = False
limiter = True

parent_path = str(pathlib.Path(__file__).parent.absolute())
source_path = parent_path + '\\json'
output_path = parent_path + '\\xlsx'
os.makedirs(source_path, exist_ok=True)
os.makedirs(output_path, exist_ok=True)

total_data_size = 0

if single_workbook:
    wb = xlsxwriter.Workbook(output_path + '\\busdata.xlsx')
    start_time = time.time()

print('-' * 150)
print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
      colored('Data extraction has started', 'magenta'))

for file_path in glob.glob(source_path + '\\*.json'):
    file_name = os.path.basename(file_path)[:-5]
    print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
          'Processing', colored(f'{file_name}.json', 'green'))

    sheet_index = 0

    with open(source_path + '\\' + file_name + '.json') as file:
        json_file = json.load(file)
        start_time = time.time()

    if not single_workbook:
        wb = xlsxwriter.Workbook(output_path + '\\' + file_name + '.xlsx')

    id_key_list = list(json_file[database_key].keys())
    sheet_length = 0
    for id_key in id_key_list:
        if id_key != 'Info':
            for date_key in list(json_file[database_key][id_key]):
                sheet_length += 1

    for id_key in id_key_list:
        if id_key == 'Info':
            continue

        date_key_list = list(json_file[database_key][id_key])

        for date_key in date_key_list:
            data_frame = pd.DataFrame.from_dict(
                json_file[database_key][id_key][date_key]).T

            time_list = data_frame.index.tolist()
            date = datetime.strptime(date_key, '%Y-%m-%d').date()
            latitude_list = data_frame['en'].tolist()
            longitude_list = data_frame['boy'].tolist()

            total_data_size += len(time_list)
            end = len(time_list) + 1
            sheet_index += 1
            sheet_name_base = date.strftime('%m-%d')
            sheet_name = f'{sheet_name_base}.{sheet_index:02}'
            ws = wb.add_worksheet(sheet_name)

            ws.set_column(0, 8, 11)

            title_format = wb.add_format()
            title_format.set_bold()
            number_format = wb.add_format()
            number_format.set_num_format('0.000')
            number_format_one_decimal = wb.add_format()
            number_format_one_decimal.set_num_format('0.0')
            time_format = wb.add_format()
            time_format.set_num_format('hh:mm:ss')
            time_format.set_align('left')
            date_format = wb.add_format()
            date_format.set_num_format('yyyy-mm-dd')
            date_format.set_align('left')

            ws.write_column('A2', time_list, time_format)
            ws.write_column('B2', [date] * len(time_list), date_format)
            ws.write_column('C2', [id_key] * len(time_list))
            ws.write_column('D2', latitude_list)
            ws.write_column('E2', longitude_list)

            title_list = ['Time', 'Date', 'ID', 'Latitude', 'Longitude',
                          'dt (s)', 'dx (m)', 'Speed (m/s)', 'Speed (km/h)', 'Time (Decimal)',
                          'Duration (hh:mm:ss)', 'Distance (km)', 'Avg. Speed (km/h)']
            for index, title in enumerate(title_list):
                ws.write(0, index, title, title_format)

            for row in range(2, len(time_list) + 1):
                ws.write_formula(
                    f'F{row + 1}', f'=IF(OR((A{row + 1}-A{row})*24*60*60 >= 30 , (A{row + 2}-A{row + 1})*24*60*60 >= 30), 0, (A{row + 1}-A{row})*24*60*60)')
                ws.write_formula(
                    f'G{row + 1}', f'=IF(OR(F{row + 1} <= 0, D{row + 1} = 0, E{row + 1} = 0, D{row} = 0, E{row} = 0), 0, IF(AND(D{row + 1}=D{row},E{row + 1}=E{row}),0,ACOS(COS(RADIANS(90-D{row + 1}))*COS(RADIANS(90-D{row}))+SIN(RADIANS(90-D{row + 1}))*SIN(RADIANS(90-D{row}))*COS(RADIANS(E{row + 1}-E{row})))*6371*1000))', cell_format=number_format)
                ws.write_formula(
                    f'H{row + 1}', f'=IF(F{row + 1} <= 0, 0, G{row + 1}/F{row + 1})', cell_format=number_format)
                if limiter:
                    ws.write_formula(
                        f'I{row + 1}', f'=MIN(80, H{row + 1}*3.6)', cell_format=number_format)
                else:
                    ws.write_formula(
                        f'I{row + 1}', f'=H{row + 1}*3.6', cell_format=number_format)
                ws.write_formula(
                    f'J{row + 1}', f'=A{row + 1}*24')

            ws.write_formula('K2', f'=SUM(F:F)/24/60/60',
                             cell_format=time_format)
            ws.write_formula('L2', f'=SUM(G:G)/1000',
                             cell_format=number_format_one_decimal)
            ws.write_formula('M2', f'=L2/K2/24',
                             cell_format=number_format_one_decimal)

            chart = wb.add_chart({'type': 'scatter',
                                  'subtype': 'straight_with_markers'})

            chart.add_series({
                'name':       'Speed (km/h)',
                'categories': f'=\'{sheet_name}\'!$A$3:$A${len(time_list) + 1}',
                'values':     f'=\'{sheet_name}\'!$I$3:$I${len(time_list) + 1}',
            })

            # min_x = int(time_list[0][:1])
            # max_x = int(time_list[-1][:1]) + 1

            # if int(time_list[0][3:4]) <= 5:
            #     min_x -= 1
            # if int(time_list[-1][3:4]) >= 55:
            #     max_x += 1

            chart.set_title({'name': f'Speed vs. Time'})
            chart.set_y_axis({'name': 'Speed', 'num_format': '0'})
            chart.set_x_axis({'name': 'Time', 'min': 0, 'max': len(time_list)})
            chart.set_legend({'none': True})
            chart.set_size({'width': 704, 'height': 400})

            ws.insert_chart('N2', chart)

            progress = int(100*sheet_index/sheet_length)
            print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
                  colored('Sheet', 'grey'), colored(f'{sheet_name}', 'blue'),
                  colored('was successfully added to the XLSX file.', 'grey'),
                  colored(f'({progress}%)', 'red'))

    if not single_workbook:
        close_workbook(wb, file_name, start_time)

if single_workbook:
    close_workbook(wb, 'busdata', start_time)        

print(f'Successfully exported a total of', colored(
    f'{total_data_size}', 'cyan'), 'data points.')
