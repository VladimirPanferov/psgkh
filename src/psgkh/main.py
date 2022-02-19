import os

from typing import NoReturn

import pandas as pd
import openpyxl as pyxl

from service.extract_data import ExtractDataService


def main(data_directory: str) -> NoReturn:
    for file in os.listdir(data_directory):
        file_path = os.path.join(data_directory, file)
        workbook = pyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active
        
        metering_device_value, accounr_and_bill\
            = ExtractDataService.get_metering_device_value_account_bill(
                sheet=sheet,
            )


if __name__ == "__main__":
    data_dir = "data/"
    main(data_dir)
