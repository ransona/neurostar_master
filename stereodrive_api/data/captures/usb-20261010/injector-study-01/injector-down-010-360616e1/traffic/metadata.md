# usb-20261010/injector-study-01/injector-down-010-360616e1/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

Native command body unavailable; exact observed USB targets are listed below.

## Discovered

- No motor target packet recorded; not evidence of a completed requested move.
- PASSIVE follow-up, despite directory name: original piston count -8572 verified idle; no reverse request.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 3,
    "last_utc": "2026-10-09T23:31:06.424791000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 3,
    "last_utc": "2026-10-09T23:31:06.455805000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 3,
    "last_utc": "2026-10-09T23:31:06.485025000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 3,
    "last_utc": "2026-10-09T23:31:06.515856000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.