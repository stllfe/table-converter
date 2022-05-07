import os

from functools import lru_cache
from pathlib import Path
from typing import Union

from contextlib import contextmanager

import win32com.client as win32
from win32com.client import constants as win32c

from . import errors
from .tables import PivotTable
from .tables import Calculation


@lru_cache(None)  # sort of singleton
def get_excel_instance(visible=False) -> object:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = visible
    return excel


def as_excel_calculation(calc: Calculation) -> object:
    try:
        calc_map = {
            Calculation.SUM: win32c.xlSum,
            Calculation.AVG: win32c.xlAverage,
            Calculation.COUNT: win32c.xlCount,
        }
        return calc_map[calc]
    except AttributeError as error:
        raise errors.ExcelNotAvailableError(error) from None
    except KeyError:
        error = NotImplementedError("Got uknown calculation: %s" % calc)
        raise errors.InternalError(error) from None


@contextmanager
def workbook(filepath: Union[str, os.PathLike], visible: bool = False) -> object:
    filepath = Path(filepath).resolve()
    excel = get_excel_instance(visible)

    try:
        workbook = excel.Workbooks.Open(filepath)
    except Exception:
        raise errors.OpenExcelError(filepath) from None
    try:
        yield workbook
        workbook.Save()
    finally:
        workbook.Close(True)
        excel.Quit()


def get_sheet(workbook: object, sheet_name: str) -> object:
    return workbook.Sheets(sheet_name)


def create_new_sheet(workbook: object, name: str) -> object:
    workbook.Sheets.Add().Name = name
    return workbook.Sheets(name)


def create_pivot_table(workbook: object, worksheet: object, table: PivotTable, show_annotations: bool = False) -> object:
    """Creates a pivot table in the given workbook from the given worksheet.

    Args:
        workbook: A workbook reference.
        worksheet: A worksheet reference.
        table: A pivot table specification 
            with all the values selected for filling the pivot tables.
        show_annotations: Whether to show columns/values annotation headers in the pivot table.
    
    Returns:
        A reference to the created pivot table worksheet.
    """

    pt = create_new_sheet(workbook, table.name)

    # pivot table location
    pt_loc = len(table.fields.filters) + 2
    
    # grab the pivot table source data
    pc = workbook.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=worksheet.UsedRange)
    
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f"'{pt.Name}'!R{pt_loc}C1", TableName=pt.Name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt.Select()
    pt.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, orientation in (
        (table.fields.filters, win32c.xlPageField), 
        (table.fields.rows, win32c.xlRowField), 
        (table.fields.columns, win32c.xlColumnField)
    ):
        for i, value in enumerate(field_list):
            pt.PivotTables(pt.Name).PivotFields(value).Orientation = orientation
            pt.PivotTables(pt.Name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for value in table.fields.values:
        calculation = as_excel_calculation(value.calculation)
        field = (
            pt.PivotTables(pt.Name)
            .AddDataField(
                pt.PivotTables(pt.Name).PivotFields(value.field), 
                value.name, calculation
            )
        )
        field.NumberFormat = value.number_format

    # Visiblity True or Valse
    pt.PivotTables(pt.Name).ShowValuesRow = show_annotations
    pt.PivotTables(pt.Name).ColumnGrand = show_annotations
