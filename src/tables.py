from enum import Enum
from typing import Sequence, NamedTuple


class Calculation(Enum):
    SUM = 'Сумма'
    AVG = 'Среднее'
    COUNT = 'Количество'


class Value(NamedTuple):
    field: str
    calculation: Calculation
    number_format: str = '0'


class Fields(NamedTuple):
    values: Sequence[Value]
    rows: Sequence[str] = tuple()
    columns: Sequence[str] = tuple()
    filters: Sequence[str] = tuple()


class PivotTable(NamedTuple):
    name: str
    fields: Fields


tables = (
    PivotTable(
        'Потребность в препаратах',
        Fields(
            rows=('МНН+Дозировка',),
            values=(
                Value('УНРЗ', Calculation.COUNT),
                Value('Потребность на год (ЕИ)', Calculation.SUM),
            )
        )
    ),
    PivotTable(
        'Схемы',
        Fields(
            rows=('Схема на УРНЗ', 'УНРЗ'),
            values=(
                Value('Потребность на год (ЕИ)', Calculation.COUNT),
            )
        )
    ),
)
