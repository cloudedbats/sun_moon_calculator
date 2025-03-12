#!/usr/bin/python3
# -*- coding:utf-8 -*-

import datetime
import utils

# MAIN.
if __name__ == "__main__":

    latitude = 56.2
    longitude = 16.4
    start_year = datetime.datetime.today().year
    end_year = start_year
    name = "ottenby"

    # latitude = 57.04432
    # longitude = 15.80785
    # start_year = 2024
    # end_year = 2025
    # name = "test"

    # Calculate.
    sun_moon_to_excel = utils.SunMoonToExcel()
    sun_moon_to_excel.generate_data(
        latitude_dd=latitude,
        longitude_dd=longitude,
        start_year=start_year,
        end_year=end_year,
    )
    # Save to file.
    years = str(start_year)
    if start_year != end_year:
        years = str(start_year) + "-" + str(end_year)
    filename = "sun_moon_" + str(years) + ".xlsx"
    if len(name) > 0:
        filename = "sun_moon_" + name + "_" + str(years) + ".xlsx"
    sun_moon_to_excel.create_report(filename)
