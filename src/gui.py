from __future__ import annotations

import abc
import logging

from typing import Iterable

from .tables import PivotTable
from .core import Params


class GUI(abc.ABC):

    def __init__(self, available_tables: Iterable[PivotTable]) -> None:
        super().__init__()
        self._available_tables = available_tables

    def __enter__(self) -> GUI:
        return self.create()

    def __exit__(self, *exc_info) -> None:
        return self.destroy()

    def handle_error(error: Exception) -> None:
        pass

    def create(self) -> GUI:
        pass

    def destroy(self) -> None:
        pass
    
    def get_params(self) -> Params:
        pass
