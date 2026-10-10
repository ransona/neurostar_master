# usb-20261010/injector-validation-04/injector-down-010-fd8a1f71/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-e9aec7d65e2148529f00990cd2acd55d",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 10
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 10 nL: delta -1614 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:42:19.564137000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:42:19.594175000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:42:19.624262000Z"
  },
  "0x70": {
    "first_raw": -6958,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 12,
    "last_utc": "2026-10-09T23:42:20.167270000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.