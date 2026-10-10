# usb-20261010/injector-validation-04/injector-up-010-7b3d325e/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-25ed45528c6c427a8c03ed06c269267d",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 10
  }
]

## Discovered

- Selector 0x70 absolute raw target -1133; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 10 nL: delta 7439 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:41:55.092919000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:41:55.124089000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:41:55.153387000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": -1133,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T23:41:54.488265000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.