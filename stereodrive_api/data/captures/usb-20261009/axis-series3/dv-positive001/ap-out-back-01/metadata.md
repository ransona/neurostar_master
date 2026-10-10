# usb-20261009/axis-series3/dv-positive001/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-18e65854e252487a88bf4e7a2718d7a0",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.05
  }
]

## Discovered

- Selector 0x50 absolute raw target 75603; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:22.564498000Z"
  },
  "0x50": {
    "first_raw": 76125,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 9,
    "last_utc": "2026-10-09T22:06:22.625398000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:22.686392000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:22.747030000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.