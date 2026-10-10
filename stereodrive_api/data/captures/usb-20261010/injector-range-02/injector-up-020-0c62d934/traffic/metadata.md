# usb-20261010/injector-range-02/injector-up-020-0c62d934/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-0e98bf6f9ff24f7f9b65d65cc109e809",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 20
  }
]

## Discovered

- Selector 0x70 absolute raw target 480; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 20 nL: delta 9052 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:16.967021000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:16.996609000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:17.026784000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": 480,
    "last_moving": 0,
    "samples": 16,
    "last_utc": "2026-10-09T23:38:16.364383000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.