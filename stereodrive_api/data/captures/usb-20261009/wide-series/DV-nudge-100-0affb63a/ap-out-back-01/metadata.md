# usb-20261009/wide-series/DV-nudge-100-0affb63a/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "nudge-usb-546905dbf0764690adae99e496c0ff35",
    "kind": "nudge",
    "axis": "DV",
    "direction": -1,
    "step_mm": 1.0
  },
  {
    "command_id": "nudge-usb-c1ede2c8173740d889340ccfffdd80a1",
    "kind": "nudge",
    "axis": "DV",
    "direction": 1,
    "step_mm": 1.0
  }
]

## Discovered

- Selector 0x60 absolute raw target 47044; profile words [1959, 8000, 8000]; tail 00010004.
- Selector 0x60 absolute raw target 41767; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:41:11.988637000Z"
  },
  "0x50": {
    "first_raw": 75864,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:41:12.049575000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T22:41:12.109576000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T22:41:12.169319000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.