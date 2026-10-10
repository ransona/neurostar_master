# usb-20261010/injector-triples-05/injector-up-100-10fec71b/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-2be8b071f65a4b7bbc707799122706dc",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 13389; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 100 nL: delta 21961 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:31.911442000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:31.941317000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:44:31.971230000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 13389,
    "last_moving": 0,
    "samples": 22,
    "last_utc": "2026-10-09T23:44:32.001446000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.