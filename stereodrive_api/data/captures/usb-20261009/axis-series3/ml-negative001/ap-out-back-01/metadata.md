# usb-20261009/axis-series3/ml-negative001/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-b1a5317a921643e3bd5733a0cbf12e1f",
    "kind": "out_and_back",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.01
  }
]

## Discovered

- Selector 0x50 absolute raw target 75551; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x50 absolute raw target 75864; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T22:03:28.256133000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 9,
    "last_utc": "2026-10-09T22:03:28.315886000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T22:03:28.375187000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T22:03:28.435053000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.