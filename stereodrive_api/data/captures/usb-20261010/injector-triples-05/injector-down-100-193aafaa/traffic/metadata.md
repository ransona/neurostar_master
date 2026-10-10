# usb-20261010/injector-triples-05/injector-down-100-193aafaa/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-243b257bea514b7a80ce69bd3ddc6c9a",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 100
  }
]

## Discovered

- Selector 0x70 absolute raw target 23701; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 100 nL: delta -21961 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:56.116135000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:56.147587000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:44:56.176927000Z"
  },
  "0x70": {
    "first_raw": 45662,
    "last_raw": 23701,
    "last_moving": 0,
    "samples": 25,
    "last_utc": "2026-10-09T23:44:56.722911000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.