# usb-20261009/wide-series/AP-nudge-100-a4297748/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-914d0c28a1e4462e882d27bccdf3af8e",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 1.0
  },
  {
    "command_id": "nudge-usb-8f35c028d06d427f8fd64fb33009546c",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 1.0
  }
]

## Discovered

- Selector 0x40 absolute raw target 111027; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105280; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 18,
    "last_utc": "2026-10-09T22:40:58.756515000Z"
  },
  "0x50": {
    "first_raw": 75864,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:58.816720000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:58.876898000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:40:58.936863000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.