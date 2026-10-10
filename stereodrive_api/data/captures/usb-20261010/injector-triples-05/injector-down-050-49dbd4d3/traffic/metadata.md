# usb-20261010/injector-triples-05/injector-down-050-49dbd4d3/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-eb27e54d78a94322b69d8f0b670d488c",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 50 nL: delta -8068 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:23.589400000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:23.619795000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:23.649968000Z"
  },
  "0x70": {
    "first_raw": -504,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T23:44:22.987935000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.