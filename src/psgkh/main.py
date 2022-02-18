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
        
        processed_data = ExtractDataService.process_data(sheet=sheet)


if __name__ == "__main__":
    data_dir = "data/"
    main(data_dir)
