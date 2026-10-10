# usb-20261010/drill-manual-01/drill.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

Native command body unavailable; exact observed USB targets are listed below.

## Discovered

- No motor target packet recorded; not evidence of a completed requested move.
- Operator-controlled drill ON/OFF writes; reported power transitions delayed. Not a direct API replay test or RPM measurement.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 41,
    "last_utc": "2026-10-09T22:48:07.578361000Z"
  },
  "0x50": {
    "first_raw": 75864,
    "last_raw": 75864,
    "last_moving": 0,
    "samples": 42,
    "last_utc": "2026-10-09T22:48:07.609006000Z"
  },
  "0x60": {
    "first_raw": 41767,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 42,
    "last_utc": "2026-10-09T22:48:07.638832000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 42,
    "last_utc": "2026-10-09T22:48:07.669268000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.