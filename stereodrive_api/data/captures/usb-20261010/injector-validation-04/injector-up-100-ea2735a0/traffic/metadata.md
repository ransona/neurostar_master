# usb-20261010/injector-validation-04/injector-up-100-ea2735a0/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-12f6831cece6407da404c5081371af5a",
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
    "last_utc": "2026-10-09T23:43:01.281996000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:43:01.312166000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:43:01.341997000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 13389,
    "last_moving": 0,
    "samples": 23,
    "last_utc": "2026-10-09T23:43:01.372589000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.