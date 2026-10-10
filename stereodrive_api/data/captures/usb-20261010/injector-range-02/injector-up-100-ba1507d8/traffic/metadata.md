# usb-20261010/injector-range-02/injector-up-100-ba1507d8/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-72a942c7e9984019a17d2252e84a6382",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 13389; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 100 nL: delta 21961 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:50.055743000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:50.086841000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:50.116101000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 13389,
    "last_moving": 0,
    "samples": 23,
    "last_utc": "2026-10-09T23:38:50.145532000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.