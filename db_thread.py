"""
Background-thread helper for non-blocking database operations.

SQLite connections created via AssetDatabase.get_connection() open and close a
brand-new connection each call, so they are safe to use from any thread.

Typical usage
-------------
    from db_thread import run_async

    def _on_search_click(self):
        self.status_label.configure(text="Searching…")
        self._set_busy(True)

        run_async(
            func=lambda: self.db.search_assets(filters),
            on_done=self._on_search_done,   # called on main Tk thread
            on_error=self._on_search_error, # called on main Tk thread
            tk_widget=self.window,
        )

    def _on_search_done(self, results):
        self._set_busy(False)
        self._populate_table(results)

    def _on_search_error(self, exc):
        self._set_busy(False)
        messagebox.showerror("Search Error", str(exc))
"""

import threading
from typing import Callable, Any, Optional


def run_async(
    func: Callable[[], Any],
    on_done: Optional[Callable[[Any], None]] = None,
    on_error: Optional[Callable[[Exception], None]] = None,
    tk_widget=None,
) -> threading.Thread:
    """Run *func* in a daemon thread; schedule callbacks back on the Tk main loop.

    Parameters
    ----------
    func:
        Callable that performs the background work (DB queries, file I/O, etc.).
        Must **not** touch any Tk widget directly.
    on_done:
        Called with the return value of *func* once it completes successfully.
        Scheduled via ``tk_widget.after(0, …)`` so it always runs on the main
        thread (safe to update widgets here).
    on_error:
        Called with the exception if *func* raises.  Same thread-safety
        guarantee as *on_done*.
    tk_widget:
        Any live Tk/CTk widget belonging to the window that initiated the call.
        Used as the scheduling anchor for ``after(0, …)``.  If ``None`` the
        callbacks are invoked directly (useful in unit tests).

    Returns
    -------
    threading.Thread
        The daemon thread that was started (already running).
    """

    def _worker():
        try:
            result = func()
            if on_done is not None:
                if tk_widget is not None:
                    # Schedule on the Tk event loop — the only thread-safe way
                    # to update Tk widgets from outside the main thread.
                    tk_widget.after(0, lambda: on_done(result))
                else:
                    on_done(result)
        except Exception as exc:  # noqa: BLE001
            if on_error is not None:
                if tk_widget is not None:
                    tk_widget.after(0, lambda: on_error(exc))
                else:
                    on_error(exc)

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    return thread
