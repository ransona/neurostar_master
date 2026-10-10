# usb-20261009/axis-series3/ml-positive005/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-701f37cccec54108a5bb97a01027b71e",
    "kind": "out_and_back",
    "axis": "ML",
    "direction": 1,
    "step_mm": 0.05
  }
]

## Discovered

- Selector 0x50 absolute raw target 76125; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 17,
    "last_utc": "2026-10-09T22:03:55.310831000Z"
  },
  "0x50": {
    "first_raw": 75864,
    "last_raw": 76125,
    "last_moving": 0,
    "samples": 22,
    "last_utc": "2026-10-09T22:03:55.371399000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 17,
    "last_utc": "2026-10-09T22:03:55.431805000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 17,
    "last_utc": "2026-10-09T22:03:55.492028000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.