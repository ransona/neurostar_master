# usb-20261009/axis-series2/recovery-ap001/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-545ab34bbc4b441e8fb0c49353258bb3",
    "kind": "nudge",
    "axis": "AP",
    "direction": 1,
    "step_mm": 0.01
  }
]

## Discovered

- Selector 0x40 absolute raw target 105280; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105228; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105332,
    "last_raw": 105228,
    "last_moving": 0,
    "samples": 8,
    "last_utc": "2026-10-09T22:01:37.417960000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:01:37.478227000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:01:37.538491000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:01:37.598650000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.