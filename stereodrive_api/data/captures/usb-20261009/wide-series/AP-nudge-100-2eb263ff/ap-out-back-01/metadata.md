# usb-20261009/wide-series/AP-nudge-100-2eb263ff/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-374ccd66288646899e2f56b6a435266b",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 1.0
  },
  {
    "command_id": "nudge-usb-0f2ab1bdf9e64e70b940e1af945ae9c1",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 1.0
  }
]

## Discovered

- Selector 0x40 absolute raw target 100055; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105802; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T22:40:00.030031000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:00.089974000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:00.121685000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:00.182507000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.