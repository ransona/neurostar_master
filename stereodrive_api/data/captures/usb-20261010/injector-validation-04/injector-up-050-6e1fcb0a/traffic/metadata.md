# usb-20261010/injector-validation-04/injector-up-050-6e1fcb0a/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-ee51e3327c4d4b2e9db241c3ba0ae732",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target 5321; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 50 nL: delta 13893 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:42:27.871174000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:27.901683000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:27.932559000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 5321,
    "last_moving": 0,
    "samples": 18,
    "last_utc": "2026-10-09T23:42:27.963175000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.