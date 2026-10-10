# usb-20261010/injector-validation-04/injector-up-100-8cb3dde5/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-1bbe4e696a0d4a13ac937de485a552ae",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 29526; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 100 nL: delta 16137 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:08.762781000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:08.792858000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:08.823688000Z"
  },
  "0x70": {
    "first_raw": 13389,
    "last_raw": 29526,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T23:43:09.367038000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.