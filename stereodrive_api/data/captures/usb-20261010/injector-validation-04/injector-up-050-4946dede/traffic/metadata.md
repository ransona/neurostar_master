# usb-20261010/injector-validation-04/injector-up-050-4946dede/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-eda8254aaa9341b295a8417ee6241b2b",
    "kind": "injector_step",
    "direction": "up",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target 13389; profile words [1820, 8000, 8000]; tail 00010004.
- Piston up 50 nL: delta 8068 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:36.410086000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:36.440688000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:36.471245000Z"
  },
  "0x70": {
    "first_raw": 5321,
    "last_raw": 13389,
    "last_moving": 0,
    "samples": 14,
    "last_utc": "2026-10-09T23:42:35.803340000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.