import json
import tempfile
from pathlib import Path
import unittest

from stereodrive_api.event_log import EventLog


class EventLogTests(unittest.TestCase):
    def test_appends_timestamped_json_records_and_syncs_each_record(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "events.jsonl"
            log = EventLog(path)
            self.assertTrue(log.write("button.click", text="Start"))
            self.assertTrue(log.write("operation.failed", error="sample failure"))

            records = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]
            self.assertEqual([record["event"] for record in records], ["button.click", "operation.failed"])
            self.assertTrue(all(record["utc"].endswith("+00:00") for record in records))
            self.assertEqual(records[1]["error"], "sample failure")


if __name__ == "__main__":
    unittest.main()
