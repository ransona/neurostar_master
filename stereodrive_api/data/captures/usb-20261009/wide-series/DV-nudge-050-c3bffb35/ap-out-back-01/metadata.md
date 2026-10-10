# usb-20261009/wide-series/DV-nudge-050-c3bffb35/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-fa258bf96eba4d7ca4c9c487dbffa06d",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-09adb6fee8c640d5b8a0b01e486c3a8b",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x60 absolute raw target 39155; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x60 absolute raw target 41819; profile words [1959, 8000, 8000]; tail 00010004.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:39:53.261328000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:39:53.321375000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 18,
    "last_utc": "2026-10-09T22:39:53.381119000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:39:53.442611000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.