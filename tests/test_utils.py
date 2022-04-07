import io
import textwrap
import pandas as pd

from src import utils


def read_md(table: str) -> pd.DataFrame:
    """Reads a markdown-style table string into a pandas DataFrame."""

    table = io.StringIO(textwrap.dedent(table).strip())
    return (
        pd.read_table(table, sep='|', header=None, index_col=None)
        .dropna(axis=1, how='all')
    )


def test_find_header():
    test_cases = [
        ["""
            |a| | | | |
            |b| | | | |
            |c| | | | |
        """, (0, 0, 0)
        ],
        ["""
            |a| | | | |
            | |b|c|d| |
            |e| | | | |
            |f| |x| | |
        """, (1, 1, 3)
        ],
        ["""
            |a| | | | |
            | |b| |d| |
            |e| | | | |
            |f| |x| | |
        """, (0, 0, 0)
        ],
        ["""
            |a| | | | |
            | |b| |d| |
            |b|c|d| | |
            |f| |x| |y|
        """, (2, 0, 2)
        ],
        ["""
            |a| | | | |
            | |b| |d| |
            |b| |d|e|f|
            |f| |x| |y|
        """, (2, 2, 4)
        ],
    ]
    for (table, expected) in test_cases:
        result = utils.find_header(read_md(table))
        assert tuple(result) == expected
