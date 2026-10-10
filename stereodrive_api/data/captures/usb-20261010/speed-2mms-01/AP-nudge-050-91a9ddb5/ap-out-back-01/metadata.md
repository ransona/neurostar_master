# usb-20261010/speed-2mms-01/AP-nudge-050-91a9ddb5/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-82de8007fa9e44b79bd1cdc2a04d842a",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.5
  },
  {
    "command_id": "nudge-usb-39716a7fd21a4fc0974fe4fe0361f148",
    "kind": "nudge",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.5
  }
]

## Discovered

- Selector 0x40 absolute raw target 99951; profile words [1306, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 103085; profile words [1306, 8000, 8000]; tail 00010104.
- Operator-controlled drill ON/OFF writes; reported power transitions delayed. Not a direct API replay test or RPM measurement.

## Final observed counts

{
  "0x40": {
    "first_raw": 102563,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 16,
    "last_utc": "2026-10-09T23:16:53.550825000Z"
  },
  "0x50": {
    "first_raw": 78581,
    "last_raw": 78581,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:16:53.611360000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:16:53.657701000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:16:53.718294000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.