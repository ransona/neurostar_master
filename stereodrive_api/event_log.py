"""Append-only, timestamped application event log."""
import json
import os
from pathlib import Path
import threading
from datetime import datetime, timezone


class EventLog:
    """Write one structured JSON record per event, flushing it to disk."""

    def __init__(self, path):
        self.path = Path(path)
        self._lock = threading.Lock()
        self.last_error = None

    def write(self, event, **fields):
        record = {
            "utc": datetime.now(timezone.utc).isoformat(),
            "event": str(event),
            **fields,
        }
        try:
            if self.last_error:
                record["previous_log_write_error"] = self.last_error
            line = json.dumps(record, ensure_ascii=False, allow_nan=False, default=str) + "\n"
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with self._lock, self.path.open("a", encoding="utf-8") as handle:
                handle.write(line)
                handle.flush()
                os.fsync(handle.fileno())
            self.last_error = None
            return True
        except Exception as exc:
            # Diagnostics must never interrupt or alter a hardware operation.
            self.last_error = f"{type(exc).__name__}: {exc}"
            return False
