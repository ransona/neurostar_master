# usb-20261009/wide-series/ML-nudge-020-f3ca9f60/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-eb3bb38e4db7433aac78bc9f328575b7",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 0.2
  },
  {
    "command_id": "nudge-usb-015afea6089f4e1d8bb1d4afce398402",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.2
  }
]

## Discovered

- Selector 0x50 absolute raw target 76909; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75603; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:27.785828000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T22:39:27.846049000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:27.906113000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:27.965748000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.