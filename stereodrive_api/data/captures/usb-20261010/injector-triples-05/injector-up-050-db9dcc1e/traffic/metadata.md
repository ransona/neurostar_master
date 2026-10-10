# usb-20261010/injector-triples-05/injector-up-050-db9dcc1e/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-b96d2152755740a185b052a0ee0d817d",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target 13389; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 50 nL: delta 8068 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:49.600633000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:49.630767000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:49.660678000Z"
  },
  "0x70": {
    "first_raw": 5321,
    "last_raw": 13389,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T23:43:49.691496000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.