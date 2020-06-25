import datetime
import numpy
from geopy import geocoders
from General.Functions import *


class Stock(object):

    def __init__(self, txt_data_str, hurricane_list, price_data_wb, start_year):

        txt_data_list = txt_data_str.split(";")
        self.price_data_wb = price_data_wb

        self.tic = txt_data_list[0].lstrip().rstrip()
        self.sec = txt_data_list[1].lstrip().rstrip()
        self.gics = txt_data_list[2].lstrip().rstrip()
        self.gics_sub = txt_data_list[3].lstrip().rstrip()
        self.loc = parse_locaction(txt_data_list[4])

        price_data_col_names = self.get_price_data()

        self.coordinates = self.get_lon_lat()
        self.price_dict = price_data_col_names[0]
        self.column_names = price_data_col_names[1]

        self.hurricane_dict = self.get_hurricane_dict(hurricane_list)

        self.filter_price_dates(start_year)

    def hurricane_dict_to_str(self):
        str_ret = ""
        for key in self.hurricane_dict.keys():
            str_ret += "Hurricane Name: " + str(key) + "\n"
            for date_dict in self.hurricane_dict[key]:
                str_ret += "\t" + str(date_dict)
                str_ret += "\t" + str(self.hurricane_dict[key][date_dict])
                str_ret += "\n"
        return str_ret

    def get_date_percent_change(self, curr_date, prev_date):    # log returns
        prev_close = self.price_dict[prev_date][self.get_column_idx("Close")]
        curr_close = self.price_dict[curr_date][self.get_column_idx("Close")]

        return numpy.log(curr_close/prev_close)

    def get_price_data(self):
        price_dict = {}
        column_names = []
        for source_sheet in self.price_data_wb.worksheets:
            if self.tic == source_sheet.title:
                column_names = [x.value.rsplit(".", 1)[-1] for x in list(source_sheet.rows)[0][1:]]
                for row in list(source_sheet.rows)[1:]:
                    date_list = [int(x) for x in row[0].value.split("-")]
                    date = datetime.date(date_list[0], date_list[1], date_list[2])
                    price_dict[date] = [x.value for x in row[1:]]
                break
        return price_dict, column_names

    def get_column_idx(self, string):
        for i, col_name in enumerate(self.column_names):
            if string.lower() in col_name.lower():
                return i
        return -1

    def get_hurricane_dict(self, hurricane_list):
        hurricane_dict = {}
        for hurricane in hurricane_list:
            overlap_dict = {}
            start_iter = 0

            prev_date = list(self.price_dict.keys())[0]
            for stock_date in self.price_dict.keys():
                for hdi, hurricane_date in enumerate(list(hurricane.data_dict.keys())[start_iter:]):
                    if stock_date == hurricane_date:
                        # print(hurricane.data_dict[hurricane_date])
                        longlat0 = self.coordinates
                        longlat1 = hurricane.data_dict[hurricane_date][1]
                        wind_speed = hurricane.data_dict[hurricane_date][2]

                        overlap_list = [stock_date,
                                        self.tic,
                                        str(self.coordinates),
                                        hurricane.name,
                                        coordinates_distance(longlat0, longlat1),
                                        self.get_date_percent_change(stock_date, prev_date),
                                        wind_speed,
                                        get_category(wind_speed)]


                        overlap_dict[stock_date] = overlap_list
                        start_iter = hdi
                        break
                prev_date = stock_date
            if overlap_dict is not {}:
                hurricane_dict[hurricane.name] = overlap_dict
        return hurricane_dict

    def get_dict_spec_col(self, col_name):
        new_dict = {}
        for key in self.price_dict.keys():
            index = self.get_column_idx(col_name)
            if index > 0:
                new_dict[key] = self.price_dict[key][index]
            else:
                new_dict[key] = -1
        return new_dict

    def get_lon_lat(self):
        for i in range(3):
            gn = geocoders.GeoNames(username="genericwsang", timeout=100)
            try:
                return [float(y) for y in
                                     [x.split(")") for x in str(gn.geocode(self.loc, exactly_one=False)).split("(")][2][
                                         0].split(",")[:2]]
            except:
                print("Could not get coordinates for " + self.loc)
                return []

    def filter_price_dates(self, start_year=None):
        new_price_dict = {}
        if start_year is not None:
            for key in self.price_dict.keys():
                if key.year >= start_year:
                    new_price_dict[key] = self.price_dict[key]
        self.price_dict = new_price_dict

    def __str__(self):
        str_ret = "Stock: " + string_length(self.tic, 4) + " | "
        str_ret += string_length(self.sec, 16, dots=True) + " | "
        str_ret += string_length(self.gics, 16, dots=True) + " | "
        str_ret += string_length(self.gics_sub, 16, dots=True) + " | "
        str_ret += string_length(self.loc, 16, dots=True) + " | "
        str_ret += str(self.coordinates)

        # if len(list(self.price_dict.keys())):
        #     str_ret += "\n"
        #     for i, key in enumerate(list(self.price_dict.keys())):
        #         if i >= 5:
        #             str_ret += "\t..."
        #             break
        #         if i != 0:
        #             str_ret += "\n"
        #         str_ret += "\tKey: " + str(key) + "\t" + str(self.price_dict[key])

        return str_ret

def parse_locaction(location):
    if "[" in location:
        for i, letter in enumerate(location):
            if "[" == letter:
                return location[0:1+1]
    return location
