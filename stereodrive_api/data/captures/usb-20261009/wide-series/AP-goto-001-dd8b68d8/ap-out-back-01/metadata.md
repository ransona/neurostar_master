# usb-20261009/wide-series/AP-goto-001-dd8b68d8/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "absolute-usb-5e9fa9bf89464248992485d54cfdd30d",
    "kind": "axis",
    "axis": "AP",
    "target_mm": 42.76,
    "method": "goto"
  }
]

## Discovered

- No motor target packet recorded; not evidence of a completed requested move.
- Historical GoTo investigation; no target packets is not arrival evidence. Inspect retained client result.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T22:31:10.897770000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T22:31:10.958230000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T22:31:11.018931000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T22:31:11.078899000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.