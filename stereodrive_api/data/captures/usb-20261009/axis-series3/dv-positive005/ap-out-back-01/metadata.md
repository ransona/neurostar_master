# usb-20261009/axis-series3/dv-positive005/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-ab7d65168b7c44a689ea84c44466ef3d",
    "kind": "out_and_back",
    "axis": "DV",
    "direction": 1,
    "step_mm": 0.02
  }
]

## Discovered

- Selector 0x60 absolute raw target 41663; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x60 absolute raw target 41819; profile words [1959, 8000, 8000]; tail 00010004.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:45.239498000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:45.299586000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 10,
    "last_utc": "2026-10-09T22:06:45.360502000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:45.419897000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.