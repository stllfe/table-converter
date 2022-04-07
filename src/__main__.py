import logging

from .core import run
from .errors import  BaseError, InternalError
from .gui import GUI
from .tables import tables


logging.basicConfig()


if __name__ == '__main__':
    log = logging.getLogger()

    with GUI(tables) as gui:
        params = gui.get_params()
        log.debug('running with params %s' % params)
        try:
            run(params)
        except Exception as error:
            if not isinstance(error, BaseError):
                log.error("unhandled internal error: '%s'" % str(error))
                error = InternalError(error)
            else:
                log.error("core error '%s' occured, trying to handle with the GUI..." % error)
            
            gui.handle_error(error)
            exit(1)
