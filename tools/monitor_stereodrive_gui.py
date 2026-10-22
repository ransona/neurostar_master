"""Timestamp StereoDrive GUI positions and its visible Avg: resistance value."""
import argparse
import csv
import time
from datetime import datetime, timezone
from pathlib import Path

from stereodrive_controller import StereoDriveController


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, default=Path('stereodrive-gui-monitor.csv'))
    parser.add_argument('--interval', type=float, default=0.1)
    parser.add_argument('--duration', type=float, default=None)
    args = parser.parse_args()
    controller = StereoDriveController()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    with args.output.open('w', newline='', encoding='utf-8') as handle:
        writer = csv.writer(handle)
        writer.writerow(['utc', 'ap_mm', 'ml_mm', 'dv_mm', 'avg_resistance'])
        try:
            while args.duration is None or time.monotonic() - started < args.duration:
                ap, ml, dv = controller.get_current_axis_position()
                avg = controller.get_resistance_avg()
                writer.writerow([datetime.now(timezone.utc).isoformat(), ap, ml, dv, avg])
                handle.flush()
                time.sleep(max(0.01, args.interval))
        except KeyboardInterrupt:
            pass
    print(f'Wrote {args.output}')


if __name__ == '__main__':
    main()
