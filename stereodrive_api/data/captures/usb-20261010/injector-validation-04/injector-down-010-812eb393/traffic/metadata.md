# usb-20261010/injector-validation-04/injector-down-010-812eb393/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-6a6e1718b08042d8992f09c676b5e27f",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 10
  }
]

## Discovered

- Selector 0x70 absolute raw target -6958; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 10 nL: delta -7438 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:12.034005000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:12.064439000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:42:10.888392000Z"
  },
  "0x70": {
    "first_raw": 480,
    "last_raw": -6958,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T23:42:10.918183000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.