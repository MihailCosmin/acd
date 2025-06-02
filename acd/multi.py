""" Functions and objects needed for multi-threading or multi-processing """

from sys import exc_info

from traceback import print_exc
from traceback import format_exc

# Qt imports
import sys

if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import QObject
    from PySide6.QtCore import Slot
    from PySide6.QtCore import Signal
    from PySide6.QtCore import QRunnable
else:
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import QObject
    from PySide2.QtCore import Slot
    from PySide2.QtCore import Signal
    from PySide2.QtCore import QRunnable


class WorkerSignals(QObject):
    """
    Defines the signals available from a running worker thread.

    Supported signals are:

    finished (no data)
        No data

    error (tuple)
        tuple (exctype, value, traceback.format_exc() )

    result (object)
        object data returned from processing, anything

    console (str)
        str message to be printed to console

    progress
        int indicating % progress

    """

    finished = Signal()
    error = Signal(tuple)
    result = Signal(object)
    console = Signal(str)
    progress = Signal(int)

class Worker(QRunnable):
    '''
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.

    :param callback: The function callback to run on this worker thread. Supplied args and
                     kwargs will be passed through to the runner.
    :type callback: function
    :param args: Arguments to pass to the callback function
    :param kwargs: Keywords to pass to the callback function

    '''

    def __init__(self,
                 fn,
                 *args, **kwargs):
        super().__init__()
    
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        
        if "result" in kwargs:
            if kwargs['result']:
                self.kwargs['result'] = self.signals.result
            else:
                del self.kwargs['result']
        if "progress" in kwargs:
            if kwargs['progress']:
                self.kwargs['progress'] = self.signals.progress
            else:
                del self.kwargs['progress']
        if "console" in kwargs:     
            if kwargs['console']:
                self.kwargs['console'] = self.signals.console
            else:
                del self.kwargs['console']

        if "error" in kwargs:  
            if kwargs['error']:
                self.kwargs['error'] = self.signals.error
            else:
                del self.kwargs['error']

    @Slot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''

        # Retrieve args/kwargs here; and fire processing using them
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            print_exc()
            exctype, value = exc_info()[:2]
            self.signals.error.emit((exctype, value, format_exc()))
        else:
            self.signals.result.emit(result)  # Return the result of the processing
        finally:
            self.signals.finished.emit()  # Done