import logging

from .core import run
from .gui import Executor, TkinterGUI
from .tables import tables


logging.basicConfig(level=logging.DEBUG)


if __name__ == '__main__':
    exc = Executor(run, tables, logger=logging.getLogger())
    gui = TkinterGUI(exc)
    gui.run()
