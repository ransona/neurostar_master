# usb-20261009/probe02/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

Native command body unavailable; exact observed USB targets are listed below.

## Discovered

- Selector 0x40 absolute raw target 105228; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105802; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105802,
    "last_moving": 0,
    "samples": 40,
    "last_utc": "2026-10-09T21:30:18.259464000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 34,
    "last_utc": "2026-10-09T21:30:18.319999000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 34,
    "last_utc": "2026-10-09T21:30:18.380170000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 34,
    "last_utc": "2026-10-09T21:30:18.440865000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.