import os
from typing import Any, Iterable, NamedTuple, Optional, Set, Tuple, Union

import pandas as pd

from src.excel import get_excel_instance

from . import errors
from .tables import PivotTable


def read_excel(path: Union[str, os.PathLike], sheet_name: str = None) -> pd.DataFrame:
    """Reads an excel file from the given path into a pandas DataFrame.

    Args:
        path: A path to read from.
        sheet_name: A specific worksheet in the file (first by default).

    Raises:
        OpenExcelFileError: If can't open the file.
        SheetNotFoundError: If can't find the specified sheet.
    """
    
    try:
        return pd.read_excel(path, sheet_name=sheet_name or 0)
    except FileNotFoundError:
        raise errors.OpenExcelError(path)
    except ValueError:
        raise errors.SheetNotFoundError(sheet_name)


def write_excel(data: pd.DataFrame, path: Union[str, os.PathLike], sheet_name: str = None) -> None:
    try:
        return data.to_excel(path, sheet_name=sheet_name)
    except Exception:
        raise errors.WriteExcelError(path)


def get_required_fields(table: PivotTable) -> Set[str]:
    fields = set()
    for source in (
        table.fields.columns, 
        table.fields.rows, (value.field for value in table.fields.values)
    ):
        for field in source:
            fields.add(field)
    return fields


class HeaderILocation(NamedTuple):
    row: int
    start_col: int
    end_col: int


def find_header(data: pd.DataFrame, search_nrows=100) -> Optional[HeaderILocation]:
    """Finds the row and cols that are more likely to be a header
    in a sparse table with lots of empty cells.

    Header in this case is the longest sequence of non-empty cells.

    Args:
        data: A sparse dataframe with lots of empty cells.
        search_nrows: How many rows to go through with a linear search.
    
    Returns:
        A tuple with the row, start column, end column indices.
    """

    def is_empty(val: Any) -> bool:
        return pd.isna(val) or not val.strip()


    def longest_nonempty_span(row: Tuple) -> Tuple[int, int]:
        n = len(row)
        last_span = ()
        for i, val in enumerate(row):
            if is_empty(val):
                continue
            j = i
            while j < n and not is_empty(row[j]):
                j += 1
            span = i, j - 1
            if width(span) > width(last_span):
                last_span = span
        return last_span


    def width(span: Tuple[int, int]) -> int:
        if not span:
            return 0
        return span[1] - span[0] + 1


    header = tuple()
    for i, row in enumerate(data.itertuples(index=False)):
        span = longest_nonempty_span(row)
        if width(span) > width(header[1:]):
            header = (i, *span)
        if i == search_nrows - 1:
            break
    
    return HeaderILocation(*header) if header else None


def shrink_to_header(data: pd.DataFrame, header: HeaderILocation) -> pd.DataFrame:
    out = data.iloc[header.row + 1:, header.start_col:header.end_col].copy()
    out.columns = data.iloc[header.row, header.start_col:header.end_col].values
    return out.reset_index(drop=True)


def remove_useless_cells(data: pd.DataFrame) -> pd.DataFrame:
    header = find_header(data)
    
    if not header:
        raise errors.HeaderNotFoundError()

    return shrink_to_header(data, header)


def validate_fields(available: Iterable[str], tables: Iterable[PivotTable]) -> None:
    for table in tables:
        required = get_required_fields(table)
        missing = set(available).difference(required)
        if missing:
            raise errors.MissingTableFieldsError(table.name, available, missing)


def fill_missing_values(data: pd.DataFrame) -> pd.DataFrame:
    return data.ffill(axis=0)


def validate_excel_available() -> None:
    try:
        get_excel_instance(visible=False)
    except Exception as error:
        raise errors.ExcelNotAvailableError(error) from None
