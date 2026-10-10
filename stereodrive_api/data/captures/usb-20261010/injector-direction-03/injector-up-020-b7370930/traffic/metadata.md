# usb-20261010/injector-direction-03/injector-up-020-b7370930/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-1fef1f61defb4330a90b90a94a0ba35f",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 20
  }
]

## Discovered

- Selector 0x70 absolute raw target 3708; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 20 nL: delta 3228 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:39:32.483035000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:39:32.514425000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:39:32.543128000Z"
  },
  "0x70": {
    "first_raw": 480,
    "last_raw": 3708,
    "last_moving": 0,
    "samples": 11,
    "last_utc": "2026-10-09T23:39:33.148681000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.