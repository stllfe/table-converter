from pathlib import Path
from typing import Iterable, Union


class BaseError(Exception):
    """The base class for all program-specific errors."""

    def __init__(self, msg: str) -> None:
        super().__init__(msg)
        self.msg = msg
        self.name = self.__class__.__name__

    def __str__(self) -> str:
        return self.msg


class InternalError(BaseError):
    """A wrapper for any other default Python exceptions."""

    def __init__(self, source: Exception) -> None:
        self.orig_name = source.__class__.__name__
        super().__init__(f'Внутренняя ошибка программы: "{source}"')


class ExcelNotAvailableError(BaseError):
    """Raised if can't run Microsoft Excel."""
    
    def __init__(self, reason: InternalError) -> None:
        self.reason = reason
        super().__init__(
            f'Не удалось запустить Microsoft Excel! {reason}\n\n'
            'Проверьте, что Excel установлен и доступен для открытия под текущим пользователем!'
        )


class OpenExcelError(BaseError):
    """Raised on I/O errors when trying to read an excel file."""

    def __init__(self, path: Union[str, Path]) -> None:
        self.path = path
        super().__init__(
            'Не удалось открыть файл Excel. ' + 
            (f'Правильно ли указан путь?\n"{path}"' if path else 'Указан пустой путь!')
        )


class SheetNotFoundError(BaseError):
    """Raised if trying to access a worksheet that is not present in excel file."""

    def __init__(self, sheet_name: str) -> None:
        self.sheet_name = sheet_name
        super().__init__(f'Указанный лист не найден: "{sheet_name}"')


class WriteExcelError(BaseError):
    """Raised on I/O errors when trying to write an excel file."""
    
    def __init__(self, path: Union[str, Path]) -> None:
        self.path = path
        super().__init__(
            'Не удалось записать файл Excel. Возможно, введен некорректный путь '
            f'или текущий пользователь не имеет достаточно разрешений.\n"{path}"'
        )


class HeaderNotFoundError(BaseError):
    """Raised if can't detect the header of the input table."""
    
    def __init__(self) -> None:
        super().__init__(
            'Невозможно найти имена колонок и заголовок файла! '
            'Проверьте, что файл отформатирован корректно.'
        )


class MissingTableFieldsError(BaseError):
    """Raised if an input table doesn't have enough fields for a pivot table."""
    
    def __init__(self, table_name: str, available: Iterable[str], missing: Iterable[str]) -> None:
        self.table_name = table_name
        self.missing = set(missing)
        self.available = set(available)
        super().__init__(
            f'Во входном файле Excel не найдены необходимые колонки для формирования отчета "{self.table_name}"!\n'
            'Доступные колонки: ' + '", "'.join(self.available) + '.\n'
            'Необходимые колонки: ' + '", "'.join(self.missing) + '.\n\n'
            'Убедитесь, что имена колонок во входном файле совпадают или укажите другой отчет!'
        )
