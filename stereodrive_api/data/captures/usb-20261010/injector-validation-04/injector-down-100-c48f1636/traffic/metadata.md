# usb-20261010/injector-validation-04/injector-down-100-c48f1636/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-296bbbdb901e49f691a096fe314e6b88",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 7564; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 100 nL: delta -21962 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:17.197287000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:17.227952000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:17.258188000Z"
  },
  "0x70": {
    "first_raw": 29526,
    "last_raw": 7564,
    "last_moving": 0,
    "samples": 23,
    "last_utc": "2026-10-09T23:43:17.803378000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.