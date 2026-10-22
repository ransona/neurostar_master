"""Fit and score a first piecewise decoder from correlated GUI/USB logs.

This is deliberately an analysis tool: it never contacts hardware. It fits
GUI Avg values against AF12 byte 5 separately for each AF12 byte-6 range.
"""
import argparse, csv, datetime, math, statistics, subprocess
from collections import Counter


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('gui_csv', type=str)
    ap.add_argument('capture', type=str)
    args = ap.parse_args()
    gui = []
    with open(args.gui_csv, encoding='utf-8') as f:
        for row in csv.DictReader(f):
            if row['avg_resistance'] not in ('', 'None'):
                gui.append((datetime.datetime.fromisoformat(row['utc']).timestamp(),
                            float(row['avg_resistance'])))
    tshark = r'C:\Program Files\Wireshark\tshark.exe'
    cmd = [tshark, '-r', args.capture, '-Y',
           'usb.bus_id == 2 && usb.device_address == 3 && usb.transfer_type == 0x03 && usb.data_len > 0',
           '-T', 'fields', '-E', 'separator=,', '-e', 'frame.time_epoch',
           '-e', 'usbcom.data.out_payload', '-e', 'usbcom.data.in_payload']
    rows = [x.split(',', 2) for x in subprocess.check_output(cmd, text=True).splitlines() if x.count(',') >= 2]
    samples = []
    for i, row in enumerate(rows):
        if row[1] != 'af12':
            continue
        reply = next((x[2] for x in rows[i + 1:] if x[2].startswith('12')), None)
        if not reply:
            continue
        t = float(row[0]); nearest = min(gui, key=lambda x: abs(x[0] - t))
        if abs(nearest[0] - t) <= 0.2:
            raw = bytes.fromhex(reply); samples.append((nearest[1], raw[5], raw[6]))
    print(f'paired samples: {len(samples)}')
    for range_byte in sorted(set(x[2] for x in samples)):
        group = [x for x in samples if x[2] == range_byte]
        if len(group) < 10:
            continue
        mx = statistics.mean(x[1] for x in group); my = statistics.mean(x[0] for x in group)
        den = sum((x[1] - mx) ** 2 for x in group)
        slope = sum((x[1] - mx) * (x[0] - my) for x in group) / den if den else 0
        intercept = my - slope * mx
        errors = [slope * x[1] + intercept - x[0] for x in group]
        corr = sum((x[1] - mx) * (x[0] - my) for x in group) / math.sqrt(den * sum((x[0] - my) ** 2 for x in group)) if den else 0
        print(f'range=0x{range_byte:02X} n={len(group)} byte5={min(x[1] for x in group)}..{max(x[1] for x in group)} '
              f'fit={slope:.4f}*byte5+{intercept:.1f} RMSE={math.sqrt(statistics.mean(e*e for e in errors)):.1f} corr={corr:.3f}')


if __name__ == '__main__':
    main()
