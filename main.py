import logging
import multiprocessing

from src.core import run
from src.gui import Executor, TkinterGUI
from src.tables import tables


logging.basicConfig(level=logging.DEBUG)


if __name__ == '__main__':
    # is needed for pyinstaller for mp to work correcly
    # https://github.com/pyinstaller/pyinstaller/wiki/Recipe-Multiprocessing  
    multiprocessing.freeze_support()

    exc = Executor(run, tables, logger=logging.getLogger())
    gui = TkinterGUI(exc)
    gui.run()
