# usb-20261010/speed-unknown-01/DV-nudge-050-bff31560/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-e616483462d641e587352928e00c155b",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-b19e0adaf4cc4d05ac7b036cf18cb8b4",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x60 absolute raw target 41714; profile words [653, 8000, 8000]; tail 00020104.
- Selector 0x60 absolute raw target 44379; profile words [653, 8000, 8000]; tail 00020004.
- Operator-controlled drill ON/OFF writes; reported power transitions delayed. Not a direct API replay test or RPM measurement.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:41.020138000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:41.080559000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 20,
    "last_utc": "2026-10-09T23:19:41.139956000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:41.199605000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.