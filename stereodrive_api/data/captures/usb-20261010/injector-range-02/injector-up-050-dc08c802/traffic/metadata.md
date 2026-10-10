# usb-20261010/injector-range-02/injector-up-050-dc08c802/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-86f30921ad85479a99c94c513c5ab2d8",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target 5321; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 50 nL: delta 13893 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:32.994004000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:33.025298000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:33.054669000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 5321,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T23:38:33.084124000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.