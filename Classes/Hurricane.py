from General.Functions import *


class Hurricane(object):

    # http://www.aoml.noaa.gov/hrd/data_sub/newHURDAT.html - Hurricane formatting

    def __init__(self, info_list, data_list_list):

        self.name = info_list[1].lstrip().rstrip()
        self.data_list_list = data_list_list
        self.data_dict = self.get_data_dict()

    def get_data_dict(self):
        data_dict = {}
        for data_list in self.data_list_list:

            # Classification
            # WV - Tropical Wave
            # TD - Tropical Depression
            # TS - Tropical Storm
            # HU - Hurricane
            # EX - Extratropical cyclone
            # SD - Subtropical depression (winds <34 kt)
            # SS - Subtropical storm (winds >34 kt)
            # LO - A low pressure system not fitting any of above descriptions
            # DB - non-tropical Disturbance not have a closed circulation"

            time_of_day = data_list[1]
            classification = data_list[3]
            coordinate_list = parse_coordinates([data_list[4], data_list[5]])
            wind_speed = int(data_list[6])       # in knots
            category = get_category(wind_speed)

            new_data_list = [time_of_day, coordinate_list, wind_speed, category]

            # if category == 4 or category == 5:
            #     print(data_list[0], self.name, new_data_list)

            data_dict[data_list[0].date()] = new_data_list
        return data_dict

    def __str__(self):
        str_ret = "Hurricane: " + string_length(self.name, 12, dots=True) + " | "
        for i, key in enumerate(self.data_dict.keys()):
            str_ret += "\n"
            str_ret += "\t" + str(key) + "\t" + str(self.data_dict[key])
            if i >= 5:
                str_ret += "\t..."
                break
        return str_ret


def parse_coordinates(lat_long):
    new_lat_long = []
    for coord in lat_long:
        str_ret = ""
        for c in coord:
            if c in "1234567890.":
                str_ret += c
        temp = float(str_ret)
        if "w" in coord.lower() or "s" in coord.lower():
            temp = temp * -1
        new_lat_long.append(temp)
    return new_lat_long


