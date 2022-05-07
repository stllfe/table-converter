import os

from pathlib import Path
from typing import Iterable, NamedTuple, Union

from . import excel, utils
from .tables import PivotTable


CLEANED_SHEET_NAME = 'Исходный лист'
EXCEL_VISIBLE = False
SHOW_PIVOT_ANNOTATIONS = False
HEADER_SEARCH_NROWS = 15


class Params(NamedTuple):
    input_path: Union[str, os.PathLike]
    output_path: Union[str, os.PathLike]
    pivot_tables: Iterable[PivotTable]
    sheet_name: str = None


def run(params: Params) -> None:
    utils.validate_excel_available()

    utils.validate_filepath(params.input_path, not_empty=True, exists=True)
    utils.validate_filepath(params.output_path, not_empty=True)

    input_path = Path(params.input_path).resolve()
    output_path = Path(params.output_path).resolve()
 
    data = utils.read_excel(input_path, params.sheet_name)
    
    data = utils.shrink_to_header(data, utils.find_header(data, HEADER_SEARCH_NROWS))
    data = utils.fill_missing_values(data)
    data = utils.add_computed_fields(data)

    for table in params.pivot_tables:
        utils.validate_fields_exist(data.columns, table)
        data = utils.cast_fields_dtypes(data, table)

    utils.write_excel(data, output_path, CLEANED_SHEET_NAME)  # should save to temp dir
    
    # open again with a native win32 API and create the pivot tables
    with excel.workbook(output_path, EXCEL_VISIBLE) as wb:
        ws = excel.get_sheet(wb, CLEANED_SHEET_NAME)
        for table in params.pivot_tables:
            excel.create_pivot_table(wb, ws, table, show_annotations=SHOW_PIVOT_ANNOTATIONS)
