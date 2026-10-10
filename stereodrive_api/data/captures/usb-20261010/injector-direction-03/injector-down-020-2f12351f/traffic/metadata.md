# usb-20261010/injector-direction-03/injector-down-020-2f12351f/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-14779f6bee86471491817a04a8d5e39d",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 20
  }
]

## Discovered

- Selector 0x70 absolute raw target -5345; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 20 nL: delta -9053 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:39:41.058563000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:39:41.088750000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:39:41.118675000Z"
  },
  "0x70": {
    "first_raw": 3708,
    "last_raw": -5345,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T23:39:40.517207000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.