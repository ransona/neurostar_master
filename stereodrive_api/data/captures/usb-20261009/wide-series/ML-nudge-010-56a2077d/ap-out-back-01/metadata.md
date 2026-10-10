# usb-20261009/wide-series/ML-nudge-010-56a2077d/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-adc7172e32f64218ab2a14d1ce1dbab4",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.1
  },
  {
    "command_id": "nudge-usb-4409e1c6613549c1963c7bdc0d4636de",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 0.1
  }
]

## Discovered

- Selector 0x50 absolute raw target 75081; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75864; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:46.396022000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T22:40:46.457284000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:46.517288000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T22:40:46.577676000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.