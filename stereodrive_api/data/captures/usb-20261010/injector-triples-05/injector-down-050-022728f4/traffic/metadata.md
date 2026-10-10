# usb-20261010/injector-triples-05/injector-down-050-022728f4/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-166c929a4ad24defb2b45b823a36cf2a",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target 7564; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 50 nL: delta -13894 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:06.559061000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:06.589414000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:06.619280000Z"
  },
  "0x70": {
    "first_raw": 21458,
    "last_raw": 7564,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T23:44:06.649674000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.