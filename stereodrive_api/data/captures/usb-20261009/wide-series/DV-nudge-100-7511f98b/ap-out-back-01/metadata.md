# usb-20261009/wide-series/DV-nudge-100-7511f98b/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-1356ef7f734246259ba08740edeb8b66",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 1.0
  },
  {
    "command_id": "nudge-usb-bd2f6aeafb7e472e97e36dd1a164f2b3",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 1.0
  }
]

## Discovered

- Selector 0x60 absolute raw target 36542; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x60 absolute raw target 41819; profile words [1959, 8000, 8000]; tail 00010004.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:34.491242000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:34.551333000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T22:40:34.611424000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:34.671470000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.