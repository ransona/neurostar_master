# usb-20261009/wide-series/AP-nudge-050-3b8a80b8/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-c077f6d1c9dd4853b7ccb660c6e42294",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-dc277841a3094fc48033c064ba047b8f",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x40 absolute raw target 102668; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105802; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 17,
    "last_utc": "2026-10-09T22:39:40.206462000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:40.265738000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:40.311793000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:40.372055000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.