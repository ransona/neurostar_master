# usb-20261009/axis-series3/dv-negative001/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-07726b7fc366484fa69694ee79609acb",
    "kind": "out_and_back",
    "axis": "DV",
    "direction": -1,
    "step_mm": 0.01
  }
]

## Discovered

- Selector 0x60 absolute raw target 41871; profile words [1959, 8000, 8000]; tail 00010004.
- Selector 0x60 absolute raw target 41767; profile words [1959, 8000, 8000]; tail 00010104.

## Final observed counts

{
  "0x40": {
    "first_raw": 105280,
    "last_raw": 105280,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T22:06:58.345489000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T22:06:58.405712000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41767,
    "last_moving": 0,
    "samples": 9,
    "last_utc": "2026-10-09T22:06:58.466749000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 5,
    "last_utc": "2026-10-09T22:06:58.512538000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.