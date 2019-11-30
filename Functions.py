import os
import openpyxl
from math import *


def remove_file(filepath):
    try:
        os.remove(filepath)
    except:
        pass


def string_length(some_var, length, dots=False):
    string = str(some_var)
    new_string = ""

    temp_dots = True
    for i in range(length):
        if len(string) > i:
            if length - i <= 3 and dots and temp_dots:
                new_string += "."
            else:
                new_string += string[i]
        elif length - i <= 3 and dots and temp_dots:
            temp_dots = False
            new_string += "."
        else:
            temp_dots = False
            new_string += " "
    return new_string


def open_xlsx(filepath):
    while True:
        try:
            return openpyxl.load_workbook(open(filepath, "rb+"))
        except:
            pass


def coordinates_distance(lat_long0, lat_long1):
    R = 6373.0

    lat1 = radians(lat_long0[0])
    lon1 = radians(lat_long0[1])
    lat2 = radians(lat_long1[0])
    lon2 = radians(lat_long1[1])

    d_lon = lon2 - lon1
    d_lat = lat2 - lat1

    a = sin(d_lat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(d_lon / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    return R * c * 0.62137


def txt_to_list(txt_path):
    txt = open(txt_path, 'r')
    txt_list = []
    for line in txt.readlines():
        if line.rstrip() != "":
            txt_list.append(line.rstrip())
    txt.close()
    return txt_list


def get_category(wind_speed):

    if 64 <= wind_speed < 82:
        return 1
    elif 83 <= wind_speed < 95:
        return 2
    elif 96 <= wind_speed < 112:
        return 3
    elif 113 <= wind_speed < 135:
        return 4
    elif wind_speed >= 137:
        return 5
    else:
        return 0
