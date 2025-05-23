import time
import statistics

class DurationStats:
    def __init__(self):
        self.durations = []

    def record(self, duration):
        self.durations.append(duration)

    def stats(self):
        if not self.durations:
            return {}
        return {
            "count": len(self.durations),
            "min": min(self.durations),
            "max": max(self.durations),
            "avg": statistics.mean(self.durations),
            "p90": statistics.quantiles(self.durations, n=10)[8],  # 90th percentile
            "p95": statistics.quantiles(self.durations, n=20)[18],  # 95th percentile
        }

duration_stats = DurationStats()

def track_duration(func):
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        duration = time.time() - start
        duration_stats.record(duration)
        return result
    return wrapper
