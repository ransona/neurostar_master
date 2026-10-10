# usb-20261009/wide-series/DV-nudge-010-2dc06eb3/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-4313eedf665646949e5b0a15793d81cb",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 0.1
  },
  {
    "command_id": "nudge-usb-bbefc627fede43c1aa152991226d0d5d",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 0.1
  }
]

## Discovered

- Selector 0x60 absolute raw target 41245; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x60 absolute raw target 41819; profile words [1959, 8000, 8000]; tail 00010004.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:49.559543000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:49.619405000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 12,
    "last_utc": "2026-10-09T22:38:49.680792000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:49.741125000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.