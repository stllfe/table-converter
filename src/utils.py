import os

from pathlib import Path
from typing import (
    Any, 
    Iterable,
    NamedTuple, 
    Optional, 
    Set, 
    Tuple, 
    Union
)

import pandas as pd

from . import errors
from .excel import get_excel_instance
from .tables import Calculation, PivotTable, Value


def read_excel(path: Union[str, os.PathLike], sheet_name: str = None, asis=False) -> pd.DataFrame:
    """Reads an excel file from the given path into a pandas DataFrame.

    Args:
        path: A path to read from.
        sheet_name: A specific worksheet in the file (first by default).
        asis: Whether to keep all values as strings or try to infer the type (default=False).

    Raises:
        OpenExcelFileError: If can't open the file.
        SheetNotFoundError: If can't find the specified sheet.
    """

    sheet_name = sheet_name or 0
    dtype = 'object' if asis else None
    try:
        return pd.read_excel(path, sheet_name=sheet_name, dtype=dtype)
    except FileNotFoundError:
        raise errors.OpenExcelError(path) from None
    except ValueError:
        raise errors.SheetNotFoundError(sheet_name) from None


def write_excel(data: pd.DataFrame, path: Union[str, os.PathLike], sheet_name: str = None) -> None:
    path = Path(path).resolve()
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        return data.to_excel(path, sheet_name=sheet_name, index=False)
    except Exception:
        raise errors.WriteExcelError(path) from None


def is_strict_numerical(value: Value) -> bool:
    return value.calculation != Calculation.COUNT


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


def find_header(data: pd.DataFrame, search_nrows=15) -> Optional[HeaderILocation]:
    """Finds the row and cols that are more likely to be a header
    in a sparse table with lots of empty cells.

    Header in this case is the longest sequence of non-empty cells.

    Args:
        data: A sparse dataframe with lots of empty cells.
        search_nrows: How many rows to go through with a linear search.

    Raises:
        HeaderNotFoundError: If no header can be found.
    
    Returns:
        A tuple with the row, start column, end column indices.
    """

    def is_empty(val: Any) -> bool:
        return pd.isna(val) or not str(val).strip()


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

    if not header:
        raise errors.HeaderNotFoundError()
    
    return HeaderILocation(*header)


def shrink_to_header(data: pd.DataFrame, header: HeaderILocation) -> pd.DataFrame:
    out = data.iloc[header.row + 1:, header.start_col:header.end_col].copy()
    out.columns = data.iloc[header.row, header.start_col:header.end_col].values
    return out.reset_index(drop=True)


def validate_fields_exist(available: Iterable[str], table: PivotTable) -> None:
    required = get_required_fields(table)
    missing = required.difference(available)
    if missing:
        raise errors.MissingTableFieldsError(table.name, available, missing)


def fill_missing_values(data: pd.DataFrame) -> pd.DataFrame:
    return data.ffill(axis=0)


def validate_excel_available() -> None:
    try:
        get_excel_instance(visible=False)
    except Exception as error:
        raise errors.ExcelNotAvailableError(error) from None


def add_computed_fields(data: pd.DataFrame) -> pd.DataFrame:
    if any(col not in data.columns for col in ('????????', '??????????????????', '??????')):
        return data
    out = data.copy()
    
    out['??????+??????????????????'] = out[['??????', '??????????????????']].apply(lambda r: ', '.join((r.??????, r.??????????????????)), axis=1)
    out['?????????? ???? ????????'] = out.groupby('????????')['??????+??????????????????'].transform(lambda x: '; '.join(x))

    last_rows = out.groupby('????????').tail(1).index
    out.loc[~out.index.isin(last_rows), '?????????? ???? ????????'] = pd.NA
    return out


def cast_fields_dtypes(data: pd.DataFrame, table: PivotTable) -> pd.DataFrame:
    fields_to_cast = set(value.field for value in table.fields.values if is_strict_numerical(value))

    for column in data:
        if column in fields_to_cast and data.dtypes[column] == 'object':
            data[column] = data[column].astype(float)
    return data


def validate_filepath(path: Union[str, os.PathLike], exists=False, not_empty=False) -> None:
    if not_empty and not path:
        raise errors.OpenExcelError(path)
    if exists and not Path(path).exists():
        raise errors.OpenExcelError(path)
    if Path(path).is_dir():
        raise errors.OpenExcelError(path)
