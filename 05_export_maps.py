import os
import folium
from folium.plugins import MousePosition, HeatMap, HeatMapWithTime, FastMarkerCluster
import webbrowser
import pandas as pd
import numpy as np
import glob
from termcolor import colored
from datetime import datetime

export_heatmapwithtime = False
export_heatmap = False
export_speedmap = True
export_markermap = False

time_key, lat_key, long_key, speed_key = 'Saat', 'Enlem', 'Boylam', 'Hız (km/h)'

path = str(os.path.dirname(os.path.realpath(__file__)))
source_path = path + '\\xlsx'
os.makedirs(source_path, exist_ok=True)

dfs = []
sheet_names = []

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
                  'Processing', colored(f'{sheet_name}', 'blue'))
            dfs.append(pd.read_excel(xls, sheet_name=sheet_name))
            sheet_names.append(sheet_name)

for i, df in enumerate(dfs):
    sheet_name = sheet_names[i]
    print(colored(datetime.now().strftime('%Y/%m/%d %H:%M:%S -'), 'cyan'),
          'Exporting', colored(f'{sheet_name}.html', 'blue'))
    time_list = np.asarray(df[time_key])[1:]
    lat_list = np.asarray(df[lat_key])[1:]
    long_list = np.asarray(df[long_key])[1:]
    speed_list = np.asarray(df[speed_key])[1:]
    location_start = [np.mean(lat_list), np.mean(long_list)]

    toner_map = True
    tiles = 'Stamen Toner' if toner_map else 'OpenStreetMap'
    zoom_start = 14
    zoom_max = 17

    if export_markermap:
        output_path = path + '\\maps\\markermaps'
        os.makedirs(output_path, exist_ok=True)

        m = folium.Map(location=location_start, tiles=tiles,
                       zoom_start=zoom_start, max_zoom=zoom_max)
        MousePosition().add_to(m)

        for i in range(len((lat_list))):
            popup_text = f'Index:{i}\nTime:{time_list[i]}\nSpeed:{speed_list[i]}'
            folium.Marker([lat_list[0], long_list[0]],
                          popup=popup_text).add_to(m)

        map_path = output_path + '\\' + file_name + '.' + sheet_name + '.html'
        m.save(map_path)

    if export_speedmap:
        output_path = path + '\\maps\\speedmaps'
        os.makedirs(output_path, exist_ok=True)

        speed_limit_list = [30, 20, 10, 0]
        color_list = ['lime', 'green', 'orange', 'red']
        gradients = []
        for i in range(len(speed_limit_list)):
            gradients.append({1: color_list[i]})

        loc_list = []
        for i in range(len(speed_limit_list)):
            loc_list.append([])

        for i, speed in enumerate(speed_list):
            for j, speed_limit in enumerate(speed_limit_list):
                if speed >= speed_limit:
                    loc_list[j].append([lat_list[i], long_list[i]])
                    break

        m = folium.Map(location=location_start, tiles=tiles,
                       zoom_start=zoom_start, max_zoom=zoom_max)
        MousePosition().add_to(m)

        for i in range(len(speed_limit_list)):
            HeatMap(data=loc_list[i], radius=4, blur=2,
                    gradient=gradients[i]).add_to(m)

        map_path = output_path + '\\' + file_name + '.' + sheet_name + '.html'
        m.save(map_path)

    if export_heatmap:
        output_path = path + '\\maps\\heatmaps'
        os.makedirs(output_path, exist_ok=True)

        m = folium.Map(location=location_start, tiles=tiles,
                       zoom_start=zoom_start, max_zoom=zoom_max)
        MousePosition().add_to(m)

        loc_list = []
        for i in range(len(lat_list)):
            loc_list.append([lat_list[i], long_list[i]])

        HeatMap(data=loc_list, radius=15, blur=10).add_to(m)

        map_path = output_path + '\\' + file_name + '.' + sheet_name + '.html'
        m.save(map_path)

    if export_heatmapwithtime:
        output_path = path + '\\maps\\heatmapswithtime'
        os.makedirs(output_path, exist_ok=True)

        m = folium.Map(location=location_start, tiles=tiles,
                       zoom_start=zoom_start, max_zoom=zoom_max)
        MousePosition().add_to(m)

        loc_list = []
        minute_list = []
        for i in range(len(time_list)):
            minute = time_list[i].strftime('%H:%M')
            if minute not in minute_list:
                minute_list.append(minute)
                loc_list.append([])

        for i in range(len(lat_list)):
            minute = time_list[i].strftime('%H:%M')
            index = minute_list.index(minute)
            loc_list[index].append([lat_list[i], long_list[i]])

        HeatMapWithTime(data=loc_list, index=minute_list).add_to(m)

        map_path = output_path + '\\' + file_name + '.' + sheet_name + '.html'
        m.save(map_path)

    # webbrowser.open(map_path)

print('-' * 100)