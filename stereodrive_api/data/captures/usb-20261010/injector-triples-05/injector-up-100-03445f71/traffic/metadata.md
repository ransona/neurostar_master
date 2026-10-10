# usb-20261010/injector-triples-05/injector-up-100-03445f71/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-2640363845614690891af0fe97df470a",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 45662; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 100 nL: delta 16136 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:47.376210000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:47.406480000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:47.436916000Z"
  },
  "0x70": {
    "first_raw": 29526,
    "last_raw": 45662,
    "last_moving": 0,
    "samples": 20,
    "last_utc": "2026-10-09T23:44:47.979305000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.