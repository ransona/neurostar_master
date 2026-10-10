# usb-20261010/injector-study-01/injector-up-010-2889d472/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "injector-usb-7aad281a4db74db598a1d39659759f46",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 10
  }
]

## Discovered

- Selector 0x70 absolute raw target -1133; profile words [1820, 8000, 8000]; tail 00010004.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T23:29:43.273474000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T23:29:43.333516000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T23:29:43.393087000Z"
  },
  "0x70": {
    "first_raw": -8572,
    "last_raw": -1133,
    "last_moving": 0,
    "samples": 10,
    "last_utc": "2026-10-09T23:29:43.457520000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.