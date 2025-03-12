#!/usr/bin/python3
# -*- coding:utf-8 -*-
# Source code: https://github.com/cloudedbats/sun_moon_calculator
# Author: Arnold Andreasson, 2025.
# License: MIT, see http://opensource.org/licenses/mit.

import datetime
from utils import sun_moon
from utils import data_to_excel


class SunMoonToExcel(object):
    """ """

    def __init__(self):
        """ """
        self.clear()
        self.configure()

    def clear(self):
        """ """
        self.columns_info = []
        self.data_rows = []

    def configure(self):
        """ """
        self.columns_info = [
            {
                "header": "Date",
                "sourceKey": "date",
                "format": "date",
                "columnWidth": 15,
            },
            {
                "header": "Latitude",
                "sourceKey": "latitude_dd",
                "format": "decimal_4",
                "columnWidth": 10,
            },
            {
                "header": "Longitude",
                "sourceKey": "longitude_dd",
                "format": "decimal_4",
                "columnWidth": 10,
            },
            {
                "header": "Sunset",
                "sourceKey": "sunset_local",
                "format": "time",
                "columnWidth": 10,
            },
            {
                "header": "Dusk",
                "sourceKey": "dusk_local",
                "format": "time",
                "columnWidth": 10,
            },
            {
                "header": "Dawn",
                "sourceKey": "dawn_local",
                "format": "time",
                "columnWidth": 10,
            },
            {
                "header": "Sunrise",
                "sourceKey": "sunrise_local",
                "format": "time",
                "columnWidth": 10,
            },
            {
                "header": "Moonrise",
                "sourceKey": "moonrise_local",
                "format": "time",
                "columnWidth": 10,
            },
            {
                "header": "Moonset",
                "sourceKey": "moonset_local",
                "format": "time",
                "columnWidth": 10,
            },
            {
                "header": "Moon phase",
                "sourceKey": "moon_phase",
                "format": "time",
                "columnWidth": 15,
            },
            {
                "header": "Moon detailed",
                "sourceKey": "moon_phase_detailed",
                "format": "text",
                "columnWidth": 20,
            },
            {
                "header": "Moon (0-28)",
                "sourceKey": "moon_phase_0to28",
                "format": "decimal",
                "columnWidth": 12,
            },
            {
                "header": "Comments",
                "sourceKey": "aggregated_comments",
                "format": "text",
                "columnWidth": 50,
            },
        ]

    def generate_data(self, latitude_dd, longitude_dd, start_year, end_year):
        """ """
        self.data_rows = []
        sun_moon_object = sun_moon.SunMoon()
        start_date = datetime.date(year=start_year, month=1, day=1)
        end_date = datetime.date(year=end_year, month=12, day=31)
        delta = datetime.timedelta(days=1)
        current_date = start_date
        while current_date <= end_date:
            #
            sun_moon_info = sun_moon_object.get_sun_moon_info(
                latitude=latitude_dd, longitude=longitude_dd, date=current_date
            )
            # Concatenate comments.
            comments = []
            for comment_key in [
                "sunset_comment",
                "dusk_comment",
                "dawn_comment",
                "sunrise_comment",
                "moonrise_comment",
                "moonset_comment",
            ]:
                comment = sun_moon_info.get(comment_key, "")
                comment = str(comment)
                if len(comment) > 0:
                    comments.append(comment)
                sun_moon_info["aggregated_comments"] = " - ".join(comments)
            #
            self.data_rows.append(sun_moon_info)
            # Next date.
            current_date += delta

    def create_report(self, file_name):
        """ """
        to_excel = data_to_excel.DataToExcel()
        to_excel.create_workbook(file_name)
        to_excel.add_worksheet(
            worksheet_name="Sun-moon",
            columns_info=self.columns_info,
            rows_dict=self.data_rows,
        )
        to_excel.close_workbook()
