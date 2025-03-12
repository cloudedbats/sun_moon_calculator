#!/usr/bin/python3
# -*- coding:utf-8 -*-
# Source code: https://github.com/cloudedbats/sun_moon_calculator
# Author: Arnold Andreasson, 2025.
# License: MIT, see http://opensource.org/licenses/mit.

import xlsxwriter


class DataToExcel:
    """ """

    def __init__(self):
        """ """
        self.clear()

    def clear(self):
        """ """
        self.workbook = None

    def create_workbook(self, report_path):
        """ """
        # Create Excel workbook.
        self.workbook = xlsxwriter.Workbook(str(report_path))
        self.define_cell_formats()

    def close_workbook(self):
        """ """
        self.workbook.close()

    def add_worksheet(self, worksheet_name, columns_info, rows_dict):
        """ """
        worksheet = self.workbook.add_worksheet(worksheet_name)
        # Header.
        headers = []
        for column_dict in columns_info:
            header = column_dict.get("header", "")
            headers.append(header)
        #
        worksheet.write_row(0, 0, headers, self.format_bold)
        # Rows.
        for row_nr, metadata_row in enumerate(rows_dict):
            for column_nr, column_dict in enumerate(columns_info):
                source_key = column_dict.get("sourceKey", "")
                value = metadata_row.get(source_key, "")
                format = column_dict.get("format", "")
                if format == "integer":
                    try:
                        value_int = int(round(float(value), 0))
                        worksheet.write_number(
                            row_nr + 1, column_nr, value_int, self.format_integer
                        )
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_integer
                        )
                elif format == "decimal":
                    try:
                        value_float = float(value)
                        worksheet.write_number(
                            row_nr + 1, column_nr, value_float, self.format_decimal
                        )
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_decimal
                        )
                elif format == "decimal_2":
                    try:
                        value_float = float(value)
                        worksheet.write_number(
                            row_nr + 1, column_nr, value_float, self.format_decimal_2
                        )
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_decimal_2
                        )
                elif format == "decimal_4":
                    try:
                        value_float = float(value)
                        worksheet.write_number(
                            row_nr + 1, column_nr, value_float, self.format_decimal_4
                        )
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_decimal_4
                        )
                elif format == "decimal_6":
                    try:
                        value_float = float(value)
                        worksheet.write_number(
                            row_nr + 1, column_nr, value_float, self.format_decimal_6
                        )
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_decimal_6
                        )
                elif format == "date":
                    try:
                        worksheet.write(row_nr + 1, column_nr, value, self.format_date)
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_date
                        )
                elif format == "time":
                    try:
                        worksheet.write(row_nr + 1, column_nr, value, self.format_time)
                    except:
                        worksheet.write_blank(
                            row_nr + 1, column_nr, "", self.format_time
                        )
                else:
                    value_str = str(value)
                    worksheet.write_string(row_nr + 1, column_nr, value_str)

        # === Adjust column width. ===
        index = 0
        for column_dict in columns_info:
            width = column_dict.get("columnWidth", 10)
            worksheet.set_column(index, index, int(width))
            index += 1

    def define_cell_formats(self):
        """ """
        self.format_bold = self.workbook.add_format({"bold": True})
        #
        self.format_bold_right = self.workbook.add_format(
            {"bold": True, "align": "right"}
        )
        #
        self.format_date = self.workbook.add_format()
        self.format_date.set_num_format("yyyy-mm-dd")
        #
        self.format_time = self.workbook.add_format()
        self.format_time.set_num_format("hh:mm:ss")
        #
        self.format_integer = self.workbook.add_format()
        self.format_integer.set_num_format("0")
        #
        self.format_decimal = self.workbook.add_format()
        self.format_decimal.set_num_format("0.0")
        #
        self.format_decimal_2 = self.workbook.add_format()
        self.format_decimal_2.set_num_format("0.00")
        #
        self.format_decimal_4 = self.workbook.add_format()
        self.format_decimal_4.set_num_format("0.0000")
        #
        self.format_decimal_6 = self.workbook.add_format()
        self.format_decimal_6.set_num_format("0.000000")
