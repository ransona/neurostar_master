# usb-20261010/injector-range-02/injector-up-010-e61e84af/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

Native command body unavailable; exact observed USB targets are listed below.

## Discovered

- No motor target packet recorded; not evidence of a completed requested move.
- Passive idle observation only; no piston movement requested.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:08.715406000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:08.745854000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:08.776405000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:08.806489000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.