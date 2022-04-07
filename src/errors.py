from pathlib import Path
from typing import Iterable, Union


class BaseError(Exception):
    """The base class for all program-specific errors."""

    def __str__(self) -> str:
        return self.__class__.__name__


class InternalError(BaseError):
    """A wrapper for any other default Python exceptions."""

    def __init__(self, source: Exception):
        self.name = source.__class__.__name__
        self.msg = str(source)

    def __str__(self) -> str:
        return "An internal error '%s': \n %s" % (self.name, self.msg)


class ExcelNotAvailableError(BaseError):
    """Raised if can't run Microsoft Excel."""
    
    def __init__(self, reason: InternalError) -> None:
        super().__init__()
        self.reason = reason


class OpenExcelError(BaseError):
    """Raised on I/O errors when trying to read an excel file."""

    def __init__(self, path: Union[str, Path]) -> None:
        super().__init__()
        self.path = path


class SheetNotFoundError(BaseError):
    """Raised if trying to access a worksheet that is not present in excel file."""

    def __init__(self, sheet_name: str) -> None:
        super().__init__()
        self.sheet_name = sheet_name


class WriteExcelError(BaseError):
    """Raised on I/O errors when trying to write an excel file."""
    
    def __init__(self, path: Union[str, Path]) -> None:
        super().__init__()
        self.path = path


class HeaderNotFoundError(BaseError):
    """Raised if can't detect the header of the input table."""
    pass


class MissingTableFieldsError(BaseError):
    """Raised if an input table doesn't have enough fields for a pivot table."""
    
    def __init__(self, table_name: str, available: Iterable[str], missing: Iterable[str]) -> None:
        super().__init__()
        self.table_name = table_name
        self.missing = set(missing)
        self.available = set(available)
