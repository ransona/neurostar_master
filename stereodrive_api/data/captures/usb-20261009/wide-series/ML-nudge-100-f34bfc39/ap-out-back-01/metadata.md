# usb-20261009/wide-series/ML-nudge-100-f34bfc39/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-65a82bbc8a7b4c4b8087b49a51bd258d",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 1.0
  },
  {
    "command_id": "nudge-usb-7e6cc80266704bd4bfcc30090f1db5e6",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 1.0
  }
]

## Discovered

- Selector 0x50 absolute raw target 70378; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75864; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:41:05.382588000Z"
  },
  "0x50": {
    "first_raw": 75864,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T22:41:05.443595000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:41:05.504020000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:41:05.565586000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.