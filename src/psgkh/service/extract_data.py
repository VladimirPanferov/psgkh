from datetime import datetime
from enum import Enum
from typing import (
    List,
    Optional,
    Tuple,
    Iterable,
)

import openpyxl as pyxl
import pandas as pd
import numpy as np


class DataType(Enum):
    TYPE_1 = "type_1"
    TYPE_2 = "type_2"
    TYPE_3 = "type_3"
    TYPE_4 = "type_4"
    TYPE_5 = "type_5"
    NOT_SUPPORT = "not support"


class ExtractDataService:
    @staticmethod
    def get_type(sheet: pyxl.Workbook) -> DataType:
        if ExtractDataService.check_type_1(sheet):
            return DataType.TYPE_1
        if ExtractDataService.check_type_2(sheet):
            return DataType.TYPE_2
        if ExtractDataService.check_type_3(sheet):
            return DataType.TYPE_3
        if False:
            return DataType.TYPE_4
        if ExtractDataService.check_type_5(sheet):
            return DataType.TYPE_5
        return DataType.NOT_SUPPORT

    @staticmethod
    def check_type_1(sheet):
        return sheet["A4"].value == "Месяц" and sheet["J6"].value == "Итого" and sheet["C4"].value

    @staticmethod
    def check_type_2(sheet):
        return sheet["A5"].value == "Месяц" and sheet["C5"].value

    @staticmethod
    def check_type_3(sheet):
        return sheet["C4"].value == "Месяц" and sheet["D4"].value

    @staticmethod
    def check_type_5(sheet):
        return sheet["A4"].value == "Месяц" and sheet["E6"].value == "Начислено" and sheet["C4"].value

    @staticmethod
    def get_cell_range(
        sheet_type: DataType
    ) -> dict:
        cell_range = {
            DataType.TYPE_1: {
                "usecols": "A:J",
                "header": 5,
                "month": "C4",
                "address": "C3",
                "column_matching": {
                    "account_number": "Лицевой счет",
                    "serial_number": "Номер прибора учета",
                    "value": "Показания",
                    "month": "Месяц начисления",
                    "address": "Адрес",
                    "room_number": "Номер квартиры",
                    "calc_value": "Начислено",
                    "credit": "Задолженность",
                    "total": "Итого",
                },
            },
            DataType.TYPE_2: {
                "usecols": "A:M",
                "header": 6,
                "month": "C5",
                "address": "C4",
                "column_matching": {
                    "account_number": "Лицевой счет",
                    "serial_number": "Номер прибора учета",
                    "value": "Показания",
                    "month": "Месяц начисления",
                    "address": "Адрес",
                    "room_number": "Номер квартиры",
                    "calc_value": "Начислено",
                    "credit": "Задолженность",
                    "total": "Итого",
                },
            },
            DataType.TYPE_3: {
                "usecols": "A:K",
                "header": 5,
                "month": "D4",
                "column_matching": {
                    "account_number": "Лицевой счет",
                    "serial_number": "Номер прибора учета",
                    "value": "Показания",
                    "month": "Месяц начисления",
                    "address": "Адрес",
                    "room_number": "Номер квартиры",
                    "calc_value": "Начислено",
                    "credit": "Задолженность",
                    "total": "Итого",
                },
            },
            DataType.TYPE_4: {
                "usecols": "A:J",
                "header": 5,
                "month": "C4",
                "address": "C3",
                "column_matching": {
                    "account_number": "Лицевой счет",
                    "serial_number": "Номер прибора учета",
                    "value": "Показание",
                    "month": "Месяц начисления",
                    "address": "Адрес",
                    "room_number": "Номер квартиры",
                    "calc_value": "Начислено",
                    "credit": "Задолженность",
                    "total": "Итого",
                },
            },
            DataType.TYPE_5: {
                "usecols": "A:E",
                "header": 5,
                "month": "C4",
                "address": "C3",
                "column_matching": {
                    "account_number": "Лицевой счет",
                    "serial_number": "Номер прибора учета",
                    "value": "Показания",
                    "month": "Месяц начисления",
                    "address": "Адрес",
                    "room_number": "Номер квартиры",
                    "calc_value": "Начислено",
                    "credit": "Задолженность",
                    "total": "Итого",
                },
            },
        }

        return cell_range[sheet_type]

    @staticmethod
    def process_data(sheet: pyxl.Workbook) -> Tuple[pd.DataFrame, pd.DataFrame]:
        sheet_type = ExtractDataService.get_type(sheet=sheet)
        metering_device_value = ExtractDataService.get_metering_device_value(
            sheet=sheet,
            sheet_type=sheet_type
        )
        account_and_bill = ExtractDataService.get_account_and_bill(
            sheet=sheet,
            sheet_type=sheet_type
        )

        return (metering_device_value, account_and_bill)

    @staticmethod
    def get_metering_device_value(sheet: pyxl.Workbook, sheet_type: DataType) -> pd.DataFrame:
        cell_range = ExtractDataService.get_cell_range(sheet_type=sheet_type)
        values = list(sheet.values)
        columns = values[cell_range["header"]]
        data = pd.DataFrame(
            values[cell_range["header"]+1:],
            columns=columns
        )
        month = sheet[cell_range["month"]].value
        metering_device_value = {
            "account_number": [],
            "serial_number": [],
            "value": [],
            "month": [],
        }

        group_keys = ExtractDataService.get_group_keys(metering_device_value)

        if sheet_type == DataType.TYPE_5:
            month = datetime.strptime(month, "%Y-%m")

        ExtractDataService.fill(
            target=metering_device_value,
            sheet_type=sheet_type,
            cell_range=cell_range,
            data=data,
            month=month,
        )

        df = pd.DataFrame(metering_device_value)
        return ExtractDataService.group_cols(
            df=df,
            group_keys=group_keys,
        )

    @staticmethod
    def fill(
        target: dict,
        sheet_type: DataType,
        cell_range: dict,
        data: pd.DataFrame,
        month: datetime,
    ):
        matching_cols = ExtractDataService.match_columns(
            target.keys(),
            cell_range["column_matching"],
        )
        account_number = None
        serial_number = None
        for row in data.iterrows():
            line = row[1]
            if sheet_type == DataType.TYPE_5:
                if line[0] == matching_cols["account_number"]:
                    account_number = line[1]
                    continue
                if line[0] == "Прибор учета":
                    serial_number = line[1]
            if ExtractDataService.check_blank_line(
                line=line,
                sheet_type=sheet_type,
            ):
                if sheet_type == DataType.TYPE_5:
                    if "credit" not in target:
                        continue
                    if line["Тариф"] == "Задолженность":
                        col_name = "credit"
                        credit = line[matching_cols["calc_value"]]
                        credit_count = len(target["account_number"]) - len(target[col_name])
                        for _ in range(credit_count):
                            target[col_name].append(credit)
                    if line["Тариф"] == "Итого":
                        col_name = "total"
                        total = line[matching_cols["calc_value"]]
                        total_count = len(target["account_number"]) - len(target[col_name])
                        for _ in range(total_count):
                            target[col_name].append(total)
                continue
            for key, value in matching_cols.items():
                if value in line:
                    target[key].append(line[value])
            if sheet_type == DataType.TYPE_5:
                target["account_number"].append(account_number)
                if "serial_number" in target:
                    target["serial_number"].append(serial_number)
        if not target["month"]:
            target["month"]\
                = [month for _ in target["account_number"]]

    @staticmethod
    def get_account_and_bill(sheet: pyxl.Workbook, sheet_type: DataType) -> pd.DataFrame:
        cell_range = ExtractDataService.get_cell_range(sheet_type=sheet_type)
        values = list(sheet.values)
        columns = values[cell_range["header"]]
        data = pd.DataFrame(
            values[cell_range["header"]+1:],
            columns=columns
        )
        month = sheet[cell_range["month"]].value
        account_and_bill = {
            "account_number": [],
            "address": [],
            "room_number": [],
            "month": [],
            "calc_value": [],
            "credit": [],
            "total": [],
        }

        group_keys = ExtractDataService.get_group_keys(account_and_bill)

        if sheet_type == DataType.TYPE_5:
            month = datetime.strptime(month, "%Y-%m")

        ExtractDataService.fill(
            target=account_and_bill,
            sheet_type=sheet_type,
            cell_range=cell_range,
            data=data,
            month=month,
        )

        if not len(account_and_bill["address"]):
            address = sheet[cell_range["address"]].value
            account_and_bill["address"] = [address for _ in account_and_bill["account_number"]]

        if sheet_type == DataType.TYPE_5:
            count = len(account_and_bill["account_number"])
            account_and_bill["room_number"] = [None for _ in range(count)]

        df = pd.DataFrame(account_and_bill)
        return ExtractDataService.group_cols(
            df=df,
            group_keys=group_keys,
        )

    @staticmethod
    def match_columns(target_cols: Iterable[str], column_matching: dict) -> dict:
        return {key: value for key, value in column_matching.items() if key in target_cols}

    @staticmethod
    def check_blank_line(line: pd.Series, sheet_type: DataType) -> bool:
        if sheet_type == DataType.TYPE_1:
            pass
        elif sheet_type == DataType.TYPE_2:
            pass
        elif sheet_type == DataType.TYPE_3:
            return line["Лицевой счет"] == "Итого"
        elif sheet_type == DataType.TYPE_4:
            pass
        elif sheet_type == DataType.TYPE_5:
            return line["Тариф"] in ("Задолженность", "Итого", "", None)\
                or line[0] in ("Лицевой счет", "")
        else:
            return False
        return False

    @staticmethod
    def group_cols(df: pd.DataFrame, group_keys: List) -> pd.DataFrame:
        df.groupby(group_keys).sum()
        return df

    @staticmethod
    def get_group_keys(target: dict):
        group_keys = list(target.keys())
        for col in ("value", "calc_value"):
            if col in group_keys:
                group_keys.remove(col)
        return group_keys
