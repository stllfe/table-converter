import os

from pathlib import Path
from typing import Iterable, NamedTuple, Union

from . import excel, utils
from .tables import PivotTable


CLEANED_SHEET_NAME = 'Исходный лист'
EXCEL_VISIBLE = False
SHOW_PIVOT_ANNOTATIONS = False


class Params(NamedTuple):
    input_path: Union[str, os.PathLike]
    output_path: Union[str, os.PathLike]
    pivot_tables: Iterable[PivotTable]
    sheet_name: str = None


def run(params: Params) -> None:
    # detect any excel errors early
    utils.validate_excel_available()

    input_path = Path(params.input_path).resolve()
    output_path = Path(params.output_path).resolve()
 
    # open a raw file and shrink unnecessary cols/rows
    data = utils.read_excel(input_path, params.sheet_name)
    data = utils.remove_useless_cells(data)

    # validate it has all the necessary columns
    utils.validate_fields(data.columns, params.pivot_tables)

    # prepare the initial data for pivoting
    data = utils.fill_missing_values(data)

    # store the result
    utils.write_excel(data, output_path, CLEANED_SHEET_NAME)
    
    # open again with a native win32 API and create the pivot tables
    with excel.workbook(output_path, EXCEL_VISIBLE) as wb:
        ws = excel.get_sheet(wb, CLEANED_SHEET_NAME)
        for table in params.pivot_tables:
            excel.create_pivot_table(wb, ws, table, show_annotations=SHOW_PIVOT_ANNOTATIONS)
