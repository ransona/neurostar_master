# usb-20261009/wide-series/AP-nudge-010-11801b40/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-bdcda3c1f50149f6b9fdcd53a5d9d3c5",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.1
  },
  {
    "command_id": "nudge-usb-1e64e13cb960492189ce9cc9b392be99",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.1
  }
]

## Discovered

- Selector 0x40 absolute raw target 104758; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105802; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 12,
    "last_utc": "2026-10-09T22:38:11.137750000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:11.197681000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:11.258337000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:11.318980000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.