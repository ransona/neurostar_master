# usb-20261010/speed-unknown-01/AP-nudge-050-16e67bd9/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-e2ce63cc46774c419df3cdc31ab6b93b",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-e70d436a97a54e65ba00bd66a32abc4c",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x40 absolute raw target 99951; profile words [653, 8000, 8000]; tail 00020104.
- Selector 0x40 absolute raw target 103085; profile words [653, 8000, 8000]; tail 00020104.
- Operator-controlled drill ON/OFF writes; reported power transitions delayed. Not a direct API replay test or RPM measurement.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 20,
    "last_utc": "2026-10-09T23:19:17.640861000Z"
  },
  "0x50": {
    "first_raw": 78581,
    "last_raw": 78581,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:17.701494000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:17.762194000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:19:17.821752000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.