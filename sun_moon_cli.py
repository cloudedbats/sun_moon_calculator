#!/usr/bin/python3
# -*- coding:utf-8 -*-
# Source code: https://github.com/cloudedbats/sun_moon_calculator
# Author: Arnold Andreasson, 2025.
# License: MIT, see http://opensource.org/licenses/mit.

import click
import datetime
import utils

current_year = datetime.date.today().year

@click.command()
@click.option(
    "--latitude",
    type=float,
    default=57.89,
    prompt="Latitude",
    help="Latitude as decimal degree, for example 57.89",
)
@click.option(
    "--longitude",
    type=float,
    default=12.34,
    prompt="Longitude",
    help="Longitude as decimal degree, for example 12.34",
)
@click.option(
    "--year",
    type=int,
    default=current_year,
    prompt="Year",
    help="Year",
)
@click.option(
    "--name",
    type=str,
    default="",
    prompt="Name (optional)",
    help="The name of the place. This field is optional.",

)
def calculate_sun_moon(latitude, longitude, year, name):
    """ """
    # Calculate.
    sun_moon_to_excel = utils.SunMoonToExcel()
    sun_moon_to_excel.generate_data(
        latitude_dd=latitude,
        longitude_dd=longitude,
        start_year=year,
        end_year=year,
    )
    # Save to file.
    filename = "sun_moon_" + str(year) + ".xlsx"
    if len(name) > 0:
        filename = "sun_moon_" + name + "_" + str(year) + ".xlsx"
    sun_moon_to_excel.create_report(filename)
 
if __name__ == "__main__":
    """ """
    print("\n\nSun moon calculator")
    print("-------------------")
    print("This tool will create an Excel file with daily information ")
    print("about sunset, dusk, dawn, sunrise, and similar for the moon.")
    print("")
    print("Source code: https://github.com/cloudedbats/sun_moon_calculator")
    print("")
    print("Input to the calculations are position and year.")
    print("Position should be given as latitude and longitude, both in")
    print("the decimal degree format with decimal point, see default values.")
    print("A name for the position can be set, but is not mandatory.")
    print("Press Ctrl-C to terminate.\n")

    try:
        calculate_sun_moon()
    except:
        pass
