# usb-20261010/injector-direction-03/injector-up-020-96fc9568/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-c096b9a2d3a44fb2a4be7551988405bd",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 20
  }
]

## Discovered

- Selector 0x70 absolute raw target 480; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 20 nL: delta 9052 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:39:24.116144000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:39:24.147651000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:39:24.176827000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 480,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T23:39:24.206830000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.