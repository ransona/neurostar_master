# usb-20261010/injector-triples-05/injector-down-100-67d13b54/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-6fca5186c5e44d9f8da8f3ee206ad784",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 7564; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 100 nL: delta -16137 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:45:04.468036000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:45:04.498217000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:45:04.528916000Z"
  },
  "0x70": {
    "first_raw": 23701,
    "last_raw": 7564,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T23:45:04.559135000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.