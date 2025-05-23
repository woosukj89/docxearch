import time

class Timer:
    def __init__(self, label=None, collector=None):
        self.label = label
        self.collector = collector  # Optional: for logging durations
        self.start = None
        self.duration = None

    def __enter__(self):
        self.start = time.perf_counter()
        return self  # So you can access .duration if needed

    def __exit__(self, exc_type, exc_value, traceback):
        self.duration = time.perf_counter() - self.start
        if self.label:
            print(f"[{self.label}] Duration: {self.duration:.6f} seconds")
        if self.collector:
            self.collector.record(self.duration)
