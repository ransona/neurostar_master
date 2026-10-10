# usb-20261009/wide-series/ML-nudge-100-06f62d08/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-6429fb922b264daa8317941ee3a101bf",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 1.0
  },
  {
    "command_id": "nudge-usb-b37355d68f1e41518f789b93a226a8fd",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 1.0
  }
]

## Discovered

- Selector 0x50 absolute raw target 81089; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75603; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:27.828725000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 20,
    "last_utc": "2026-10-09T22:40:27.890029000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:27.950299000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:28.010529000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.