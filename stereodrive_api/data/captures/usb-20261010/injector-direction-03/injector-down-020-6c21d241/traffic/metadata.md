# usb-20261010/injector-direction-03/injector-down-020-6c21d241/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-2543ad23173d4273875654624e8a889d",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 20
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 20 nL: delta -3227 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:39:49.466431000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:39:49.497475000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:39:49.526648000Z"
  },
  "0x70": {
    "first_raw": -5345,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 12,
    "last_utc": "2026-10-09T23:39:49.556655000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.