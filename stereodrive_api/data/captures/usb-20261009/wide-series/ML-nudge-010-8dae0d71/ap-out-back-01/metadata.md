# usb-20261009/wide-series/ML-nudge-010-8dae0d71/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-25c34b2f41324fe39f9484d7dfa16bb8",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 0.1
  },
  {
    "command_id": "nudge-usb-8ae1505417da4436985b381110d1b390",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.1
  }
]

## Discovered

- Selector 0x50 absolute raw target 76386; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75603; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:43.926073000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T22:38:43.985751000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:44.045847000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:44.106124000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.