# usb-20261009/wide-series/AP-nudge-020-db132baf/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-1acbb21538eb4c92a258ecdf2e8e4ec9",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.2
  },
  {
    "command_id": "nudge-usb-869606e4102f446ca3a2ab1dda259270",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.2
  }
]

## Discovered

- Selector 0x40 absolute raw target 104235; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105802; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T22:38:55.664130000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:55.723830000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:55.784496000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:38:55.846143000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.