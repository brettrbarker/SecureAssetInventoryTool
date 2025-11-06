"""
Performance monitoring and optimization utilities.
Tracks operation timing and provides performance insights.
"""

import time
import functools
from typing import Dict, List, Any, Callable
from datetime import datetime
import threading

class PerformanceTimer:
    """Context manager for timing operations."""
    
    def __init__(self, operation_name: str, logger=None):
        self.operation_name = operation_name
        self.logger = logger
        self.start_time = None
        self.end_time = None
    
    def __enter__(self):
        self.start_time = time.time()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_time = time.time()
        duration = self.end_time - self.start_time
        
        if self.logger:
            self.logger.info(f"Operation '{self.operation_name}' completed in {duration:.3f} seconds")
        
        # Add to global performance tracker
        PerformanceTracker.instance().add_timing(self.operation_name, duration)
    
    @property
    def duration(self) -> float:
        """Get the duration of the operation."""
        if self.start_time and self.end_time:
            return self.end_time - self.start_time
        return 0.0

def performance_monitor(operation_name: str = None):
    """Decorator to monitor function performance."""
    def decorator(func: Callable) -> Callable:
        name = operation_name or f"{func.__module__}.{func.__name__}"
        
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            with PerformanceTimer(name):
                return func(*args, **kwargs)
        return wrapper
    return decorator

class PerformanceTracker:
    """Singleton class to track performance metrics across the application."""
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        # Prevent re-initialization
        if hasattr(self, 'timings'):
            return
            
        self.timings: Dict[str, List[float]] = {}
        self.operation_counts: Dict[str, int] = {}
        self.slow_operations: List[Dict[str, Any]] = []
        self.slow_threshold = 2.0  # seconds
    
    @classmethod
    def instance(cls):
        """Get the singleton instance."""
        return cls()
    
    def add_timing(self, operation: str, duration: float):
        """Add a timing measurement."""
        if operation not in self.timings:
            self.timings[operation] = []
            self.operation_counts[operation] = 0
        
        self.timings[operation].append(duration)
        self.operation_counts[operation] += 1
        
        # Track slow operations
        if duration > self.slow_threshold:
            self.slow_operations.append({
                'operation': operation,
                'duration': duration,
                'timestamp': datetime.now()
            })
            
            # Keep only recent slow operations (last 100)
            if len(self.slow_operations) > 100:
                self.slow_operations = self.slow_operations[-100:]
    
    def get_stats(self, operation: str = None) -> Dict[str, Any]:
        """Get performance statistics."""
        if operation:
            return self._get_operation_stats(operation)
        else:
            return self._get_all_stats()
    
    def _get_operation_stats(self, operation: str) -> Dict[str, Any]:
        """Get statistics for a specific operation."""
        if operation not in self.timings:
            return {}
        
        timings = self.timings[operation]
        return {
            'operation': operation,
            'count': len(timings),
            'total_time': sum(timings),
            'average_time': sum(timings) / len(timings),
            'min_time': min(timings),
            'max_time': max(timings),
            'last_execution': timings[-1] if timings else 0
        }
    
    def _get_all_stats(self) -> Dict[str, Any]:
        """Get statistics for all operations."""
        stats = {}
        for operation in self.timings:
            stats[operation] = self._get_operation_stats(operation)
        
        return {
            'operations': stats,
            'total_operations': sum(self.operation_counts.values()),
            'slow_operations_count': len(self.slow_operations),
            'recent_slow_operations': self.slow_operations[-10:] if self.slow_operations else []
        }
    
    def get_performance_report(self) -> str:
        """Generate a formatted performance report."""
        stats = self.get_stats()
        
        report = ["=== Performance Report ===\n"]
        report.append(f"Total Operations: {stats['total_operations']}")
        report.append(f"Slow Operations: {stats['slow_operations_count']}")
        report.append("")
        
        # Top 10 slowest operations by average time
        operations = list(stats['operations'].values())
        operations.sort(key=lambda x: x.get('average_time', 0), reverse=True)
        
        report.append("Slowest Operations (by average time):")
        for i, op in enumerate(operations[:10], 1):
            report.append(f"{i:2d}. {op['operation']}: {op['average_time']:.3f}s avg "
                         f"({op['count']} executions)")
        
        report.append("")
        
        # Recent slow operations
        if stats['recent_slow_operations']:
            report.append("Recent Slow Operations:")
            for slow_op in stats['recent_slow_operations']:
                timestamp = slow_op['timestamp'].strftime('%H:%M:%S')
                report.append(f"  {timestamp}: {slow_op['operation']} took {slow_op['duration']:.3f}s")
        
        return "\n".join(report)
    
    def reset_stats(self):
        """Reset all performance statistics."""
        self.timings.clear()
        self.operation_counts.clear()
        self.slow_operations.clear()

class DatabasePerformanceOptimizer:
    """Optimizes database operations for better performance."""
    
    def __init__(self, database_service):
        self.database_service = database_service
        self.query_cache = {}
        self.cache_ttl = 300  # 5 minutes
        self.cache_timestamps = {}
    
    def get_cached_or_execute(self, cache_key: str, query_func: Callable, *args, **kwargs):
        """Get cached result or execute query and cache result."""
        current_time = time.time()
        
        # Check if we have a valid cached result
        if (cache_key in self.query_cache and 
            cache_key in self.cache_timestamps and
            current_time - self.cache_timestamps[cache_key] < self.cache_ttl):
            return self.query_cache[cache_key]
        
        # Execute query and cache result
        with PerformanceTimer(f"DB Query: {cache_key}"):
            result = query_func(*args, **kwargs)
        
        self.query_cache[cache_key] = result
        self.cache_timestamps[cache_key] = current_time
        
        return result
    
    def clear_cache(self):
        """Clear the query cache."""
        self.query_cache.clear()
        self.cache_timestamps.clear()
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        current_time = time.time()
        valid_entries = sum(1 for ts in self.cache_timestamps.values() 
                          if current_time - ts < self.cache_ttl)
        
        return {
            'total_cached_queries': len(self.query_cache),
            'valid_cached_queries': valid_entries,
            'cache_hit_potential': f"{(valid_entries / len(self.query_cache) * 100):.1f}%" if self.query_cache else "0%"
        }

# Global performance tracker
performance_tracker = PerformanceTracker()
