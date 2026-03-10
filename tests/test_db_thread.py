"""
Tests for db_thread.run_async.

Because run_async uses tk_widget.after(0, callback) when a widget is supplied,
these tests pass tk_widget=None to invoke callbacks synchronously
(as documented in the module's own docstring).
"""

import sys
import os
import threading
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from db_thread import run_async


class TestRunAsync(unittest.TestCase):

    # ---- happy path --------------------------------------------------------

    def test_on_done_called_with_result(self):
        results = []
        t = run_async(lambda: 42, on_done=results.append, tk_widget=None)
        t.join(timeout=2)
        self.assertEqual(results, [42])

    def test_on_done_receives_none_result(self):
        results = []
        run_async(lambda: None, on_done=results.append, tk_widget=None).join(timeout=2)
        self.assertEqual(results, [None])

    def test_return_value_is_thread(self):
        t = run_async(lambda: 1, tk_widget=None)
        self.assertIsInstance(t, threading.Thread)
        t.join(timeout=2)

    def test_thread_is_daemon(self):
        t = run_async(lambda: 1, tk_widget=None)
        self.assertTrue(t.daemon)
        t.join(timeout=2)

    def test_func_executes_in_background(self):
        """The function must run off the calling thread."""
        thread_ids = []
        calling_thread_id = threading.get_ident()

        def _work():
            thread_ids.append(threading.get_ident())
            return True

        t = run_async(_work, tk_widget=None)
        t.join(timeout=2)
        self.assertEqual(len(thread_ids), 1)
        self.assertNotEqual(thread_ids[0], calling_thread_id)

    # ---- error path --------------------------------------------------------

    def test_on_error_called_on_exception(self):
        errors = []

        def _bad():
            raise ValueError("test error")

        run_async(_bad, on_error=errors.append, tk_widget=None).join(timeout=2)
        self.assertEqual(len(errors), 1)
        self.assertIsInstance(errors[0], ValueError)
        self.assertIn("test error", str(errors[0]))

    def test_on_done_not_called_on_exception(self):
        done_calls = []
        errors = []

        run_async(
            lambda: (_ for _ in ()).throw(RuntimeError("boom")),
            on_done=done_calls.append,
            on_error=errors.append,
            tk_widget=None,
        ).join(timeout=2)
        self.assertEqual(len(done_calls), 0)
        self.assertEqual(len(errors), 1)

    def test_no_callbacks_does_not_raise(self):
        """Providing no callbacks must silently complete."""
        t = run_async(lambda: 1, tk_widget=None)
        t.join(timeout=2)  # Should not raise

    def test_on_error_not_called_on_success(self):
        errors = []
        run_async(lambda: 99, on_error=errors.append, tk_widget=None).join(timeout=2)
        self.assertEqual(len(errors), 0)

    # ---- concurrency -------------------------------------------------------

    def test_multiple_calls_run_concurrently(self):
        """Two run_async calls should overlap (not block each other)."""
        log = []
        event = threading.Event()

        def slow():
            event.wait(timeout=2)
            log.append("slow_done")
            return "slow"

        def fast():
            log.append("fast_done")
            return "fast"

        t_slow = run_async(slow, tk_widget=None)
        t_fast = run_async(fast, tk_widget=None)

        t_fast.join(timeout=2)
        self.assertIn("fast_done", log)
        # slow hasn't finished yet (event not set)
        self.assertNotIn("slow_done", log)

        event.set()
        t_slow.join(timeout=2)
        self.assertIn("slow_done", log)

    def test_on_done_receives_complex_return_value(self):
        results = []
        expected = {"key": [1, 2, 3], "val": "hello"}
        run_async(lambda: expected, on_done=results.append, tk_widget=None).join(timeout=2)
        self.assertEqual(results[0], expected)


if __name__ == "__main__":
    unittest.main()
