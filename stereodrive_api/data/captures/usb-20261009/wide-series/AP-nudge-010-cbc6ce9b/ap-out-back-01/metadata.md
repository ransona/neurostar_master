# usb-20261009/wide-series/AP-nudge-010-cbc6ce9b/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-5e3ea0b47d0947d2a68e342cbc8bdb19",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.1
  },
  {
    "command_id": "nudge-usb-a30fe09c6e2c4f668121e335555cb61d",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.1
  }
]

## Discovered

- Selector 0x40 absolute raw target 106325; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105280; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 10,
    "last_utc": "2026-10-09T22:40:40.142887000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:40.204653000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:40.265590000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:40.325798000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.