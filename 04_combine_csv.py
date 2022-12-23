import os
import glob
import pathlib
import pandas as pd
from termcolor import colored
from datetime import datetime

combined_csv_name = 'combined'

time_key, lat_key, long_key, speed_key = 'Saat', 'Enlem', 'Boylam', 'Hız (km/h)'
time_key_new, lat_key_new, long_key_new, speed_key_new = 'Time', 'Latitude', 'Longitude', 'Speed'

parent_path = str(pathlib.Path(__file__).parent.absolute())
source_path = parent_path + '\\csv'
output_path = parent_path + '\\csv'
os.makedirs(source_path, exist_ok=True)
os.makedirs(output_path, exist_ok=True)

total_data_size = 0
dfs = []
# df_combined = pd.DataFrame()

print('-' * 100)
print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
      colored('Data extraction has started', 'magenta'))

for file_path in glob.glob(source_path + '\\*.csv'):
    file_name = os.path.basename(file_path)[:-4]
    if file_name == combined_csv_name:
        continue

    print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
          'Processing', colored(f'{file_name}.csv', 'green'))

    df = pd.read_csv(file_path)
    dfs.append(df)
#     df_combined = df_combined.append(df, ignore_index = True)

    total_data_size += len(df)

csv_path = output_path + '\\' + combined_csv_name + '.csv'

df_combined = pd.concat(dfs)
df_combined.to_csv(csv_path, index=False)

print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
      colored('Data extraction has ended', 'magenta'))
print(f'Processed {total_data_size} data points')
print('-' * 100)