# usb-20261009/wide-series/AP-goto-001-50e3f3ac/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "absolute-usb-16aee5e2fe314b3dbc026b61a704549c",
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
    "samples": 29,
    "last_utc": "2026-10-09T22:37:25.070592000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 29,
    "last_utc": "2026-10-09T22:37:25.131707000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 29,
    "last_utc": "2026-10-09T22:37:25.191769000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 29,
    "last_utc": "2026-10-09T22:37:25.252080000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.