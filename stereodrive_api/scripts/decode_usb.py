"""Passive decoder for completed, scoped StereoDrive USBPcap captures.

Reads files through TShark -r only. Never captures, opens a serial device, or
sends controller packets. Unknown payloads and unverified tails remain opaque.
"""
import argparse
import csv
import json
import subprocess
from collections import Counter
from datetime import datetime
from pathlib import Path

TSHARK = r'C:\Program Files\Wireshark\tshark.exe'
AXES = {0x40: 'AP', 0x50: 'ML', 0x60: 'DV', 0x70: 'PISTON'}

def decode(payload, direction):
    b = bytes.fromhex(payload.replace(':', ''))
    result = {'kind': 'opaque', 'length': len(b), 'hex': b.hex()}
    if direction == 'OUT' and b[:1] == b'\xaf':
        result['opcode'] = b[1] if len(b) > 1 else None
        if len(b) == 17 and b[1] == 0x0c:
            result.update(kind='target', selector=b[2], axis=AXES.get(b[2], 'unverified'),
                          target_raw=int.from_bytes(b[3:7], 'little', signed=True),
                          tail_hex=b[7:].hex())
        elif len(b) == 3 and b[1] in (0x0e, 0x0f):
            result.update(kind='poll' if b[1] == 0x0e else 'stop', selector=b[2],
                          axis=AXES.get(b[2], 'unverified'))
    elif direction == 'IN' and b:
        result['opcode'] = b[0]
        if b[0] == 0x0c and len(b) == 9:
            result.update(kind='target_ack', raw_field_1=int.from_bytes(b[1:5], 'little', signed=True),
                          raw_field_2=int.from_bytes(b[5:9], 'little', signed=True))
        elif b[0] == 0x0e and len(b) >= 11:
            result.update(kind='status', device_clock=int.from_bytes(b[1:5], 'little'),
                          selector=b[5], axis=AXES.get(b[5], 'unverified'),
                          position_raw=int.from_bytes(b[6:10], 'little', signed=True),
                          observed_motion_byte=b[10], tail_hex=b[11:].hex())
    return result

def read_capture(path):
    output = subprocess.run([TSHARK, '-r', str(path), '-Y', 'usbcom', '-T', 'json'],
                            check=True, capture_output=True, text=True, encoding='utf-8').stdout
    rows = []
    for packet in json.loads(output):
        layers = packet['_source']['layers']
        usb, frame, serial = layers['usb'], layers['frame'], layers.get('usbcom', {})
        for direction, field in [('OUT', 'usbcom.data.out_payload'), ('IN', 'usbcom.data.in_payload')]:
            payload = serial.get(field)
            if not payload:
                continue
            row = dict(capture=path.parent.name, frame=int(frame['frame.number']),
                       utc=frame['frame.time_utc'], bus=int(usb['usb.bus_id']),
                       device=int(usb['usb.device_address']), endpoint=usb['usb.endpoint_address'],
                       usb_status=usb['usb.usbd_status'], direction=direction)
            row.update(decode(payload, direction))
            rows.append(row)
    return rows

def summarize(rows):
    targets = [r for r in rows if r['kind'] == 'target']
    acks = [r for r in rows if r['kind'] == 'target_ack']
    moves = []
    for index, target in enumerate(targets):
        end = targets[index + 1]['frame'] if index + 1 < len(targets) else float('inf')
        ack = next((r for r in acks if target['frame'] < r['frame'] < end), None)
        arrival = next((r for r in rows if target['frame'] < r['frame'] < end
                        and r['kind'] == 'status' and r['selector'] == target['selector']
                        and r['position_raw'] == target['target_raw'] and r['observed_motion_byte'] == 0), None)
        moves.append(dict(target=target, acknowledgement=ack, first_idle_at_target=arrival))
    return dict(serial_transfers=len(rows), devices=dict(Counter(r['device'] for r in rows)),
                kinds=dict(Counter(r['kind'] for r in rows)), moves=moves)

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('captures', nargs='+', type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    summaries = {}
    for capture in args.captures:
        if not capture.is_file():
            raise ValueError(f'Missing completed capture: {capture}')
        rows = read_capture(capture)
        summaries[capture.parent.name] = summarize(rows)
        (args.output / (capture.parent.name + '.json')).write_text(json.dumps(rows, indent=2), encoding='utf-8')
        fields = sorted(set().union(*(r.keys() for r in rows)))
        with (args.output / (capture.parent.name + '.csv')).open('w', newline='', encoding='utf-8') as output:
            writer = csv.DictWriter(output, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)
    (args.output / 'summary.json').write_text(json.dumps(summaries, indent=2), encoding='utf-8')
    for name, summary in summaries.items():
        print(name, [(m['target']['frame'], m['target']['axis'], m['target']['target_raw'],
                      None if m['acknowledgement'] is None else m['acknowledgement']['raw_field_2'],
                      None if m['first_idle_at_target'] is None else m['first_idle_at_target']['frame'])
                     for m in summary['moves']])

if __name__ == '__main__':
    main()
