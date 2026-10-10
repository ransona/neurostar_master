# usb-20261009/wide-series/DV-nudge-020-969108ba/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-b37a250f3c8840cb84d59ddb1ae3710a",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 0.2
  },
  {
    "command_id": "nudge-usb-e9bdf50515bb402d89934d722f6fef47",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 0.2
  }
]

## Discovered

- Selector 0x60 absolute raw target 40722; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x60 absolute raw target 41819; profile words [1959, 8000, 8000]; tail 00010004.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:39:33.658859000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:39:33.718552000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 13,
    "last_utc": "2026-10-09T22:39:33.778344000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:39:33.839427000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.