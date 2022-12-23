import os
import glob
import pathlib
import pandas as pd
import numpy as np
from termcolor import colored
from datetime import datetime

time_key, lat_key, long_key, speed_key = 'Saat', 'Enlem', 'Boylam', 'Hız (km/h)'
time_key_new, lat_key_new, long_key_new, speed_key_new = 'Time', 'Latitude', 'Longitude', 'Speed'

parent_path = str(pathlib.Path(__file__).parent.absolute())
source_path = parent_path + '\\xlsx'
output_path = parent_path + '\\csv'
os.makedirs(source_path, exist_ok=True)
os.makedirs(output_path, exist_ok=True)

total_data_size = 0

print('-' * 100)
print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
      colored('Data extraction has started', 'magenta'))

for file_path in glob.glob(source_path + '\\*.xlsx'):
    file_name = os.path.basename(file_path)[:-5]
    print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
          'Processing', colored(f'{file_name}.xlsx', 'green'))

    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        if sheet_name != 'Index':
            print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
                  'Exporting', colored(f'{sheet_name}.csv', 'blue'))
            df = pd.read_excel(xls, sheet_name=sheet_name)

            time_list = np.asarray(df[time_key])[1:]
            lat_list = np.asarray(df[lat_key])[1:]
            long_list = np.asarray(df[long_key])[1:]
            speed_list = np.asarray(df[speed_key])[1:]

            for i in range(len(time_list)):
                time_list[i] = round((time_list[i].hour + time_list[i].minute / 60 + time_list[i].second / 3600)/24, 5)
                speed_list[i] = round(speed_list[i], 5)

            data = {time_key_new: time_list,
                    lat_key_new: lat_list,
                    long_key_new: long_list,
                    speed_key_new: speed_list}

            df_new = pd.DataFrame(data)

            csv_path = output_path + '\\' + sheet_name + '.csv'
            df_new.to_csv(csv_path, index=False)
            total_data_size += len(time_list)

print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
      colored('Data extraction has ended', 'magenta'))
print(f'Processed {total_data_size} data points')
print('-' * 100)