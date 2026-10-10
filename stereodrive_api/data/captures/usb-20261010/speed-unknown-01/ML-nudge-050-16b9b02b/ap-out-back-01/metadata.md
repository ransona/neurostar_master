# usb-20261010/speed-unknown-01/ML-nudge-050-16b9b02b/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-1a5c316594bf4ebbb72bac3cc15c5146",
    "kind": "nudge",
    "axis": "ML",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-3a27ce3f49ae4d0f9a4d265e882b3193",
    "kind": "nudge",
    "axis": "ML",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x50 absolute raw target 81193; profile words [653, 8000, 8000]; tail 00020104.
- Selector 0x50 absolute raw target 78320; profile words [653, 8000, 8000]; tail 00020104.
- Operator-controlled drill ON/OFF writes; reported power transitions delayed. Not a direct API replay test or RPM measurement.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:30.697909000Z"
  },
  "0x50": {
    "first_raw": 78581,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T23:19:30.757965000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:30.817831000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:30.877795000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.