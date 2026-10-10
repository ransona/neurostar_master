# usb-20261010/injector-triples-05/injector-up-050-bf3c51c9/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-1b9ea16150044287988ff0f6038fff52",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target 21458; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 50 nL: delta 8069 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:57.933659000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:57.963840000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:43:57.994492000Z"
  },
  "0x70": {
    "first_raw": 13389,
    "last_raw": 21458,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T23:43:58.025240000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.