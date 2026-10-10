# usb-20261010/injector-validation-04/injector-down-100-298e5f56/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-a0025a389dbf48a8a7faf6ee6a0c7b4a",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 100 nL: delta -16136 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:43:25.575228000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:43:25.605585000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:43:25.634993000Z"
  },
  "0x70": {
    "first_raw": 7564,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T23:43:25.665892000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.