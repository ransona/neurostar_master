# usb-20261010/injector-range-02/injector-down-020-a45552c2/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-dc95b33ef77e4819b2a79fa6f016e0ed",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 20
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 20 nL: delta -9052 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:24.372326000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:24.403708000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 6,
    "last_utc": "2026-10-09T23:38:24.433227000Z"
  },
  "0x70": {
    "first_raw": 480,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 15,
    "last_utc": "2026-10-09T23:38:24.464535000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.