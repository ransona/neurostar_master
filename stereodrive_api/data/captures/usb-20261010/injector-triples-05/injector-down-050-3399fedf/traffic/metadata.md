# usb-20261010/injector-triples-05/injector-down-050-3399fedf/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-54e4e549a5a54190b96556ccec5024ac",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target -504; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 50 nL: delta -8068 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:15.038275000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:15.068656000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:15.099367000Z"
  },
  "0x70": {
    "first_raw": 7564,
    "last_raw": -504,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T23:44:15.642970000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.