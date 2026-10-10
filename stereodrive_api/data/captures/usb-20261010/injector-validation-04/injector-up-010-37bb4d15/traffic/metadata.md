# usb-20261010/injector-validation-04/injector-up-010-37bb4d15/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-17b9a1beb8f64aa8962b2ca21403dae6",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 10
  }
]

## Discovered

- Selector 0x70 absolute raw target 480; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 10 nL: delta 1613 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:03.428176000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:03.458392000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:03.489437000Z"
  },
  "0x70": {
    "first_raw": -1133,
    "last_raw": 480,
    "last_moving": 0,
    "samples": 10,
    "last_utc": "2026-10-09T23:42:03.520176000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.