import itertools
import threading

class UniqueCounter:
    _lock = threading.Lock()
    _counter = itertools.count()

    @classmethod
    def next(cls):
        with cls._lock:
            return next(cls._counter)
