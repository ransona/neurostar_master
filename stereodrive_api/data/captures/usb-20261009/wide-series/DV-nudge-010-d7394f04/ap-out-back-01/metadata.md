# usb-20261009/wide-series/DV-nudge-010-d7394f04/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-d5d1e5b508f3429d8b989c8f7787f8c9",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 0.1
  },
  {
    "command_id": "nudge-usb-7db1b0e649434403ba931c95dcacabaf",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 0.1
  }
]

## Discovered

- Selector 0x60 absolute raw target 42341; profile words [1959, 8000, 8000]; tail 00010004.
- Selector 0x60 absolute raw target 41767; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:52.258160000Z"
  },
  "0x50": {
    "first_raw": 75864,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:52.318322000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 12,
    "last_utc": "2026-10-09T22:40:52.378854000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:52.439046000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.