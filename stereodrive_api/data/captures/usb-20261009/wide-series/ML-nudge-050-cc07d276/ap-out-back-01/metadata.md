# usb-20261009/wide-series/ML-nudge-050-cc07d276/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-1cfd6ea52ba54ddfabf28acbc4f0a97e",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-bdfc7db5e32345cea4ed2fc3841b526f",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x50 absolute raw target 78476; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75603; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:46.956297000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T22:39:47.016833000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:47.077155000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:39:47.137420000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.