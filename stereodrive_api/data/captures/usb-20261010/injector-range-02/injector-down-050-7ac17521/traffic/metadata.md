# usb-20261010/injector-range-02/injector-down-050-7ac17521/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-f5a6b94ecd204b53bf66fd1f143a3bb3",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 50 nL: delta -13893 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:41.777387000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:41.807561000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:38:41.837390000Z"
  },
  "0x70": {
    "first_raw": 5321,
    "last_raw": -8572,
    "last_moving": 0,
    "samples": 19,
    "last_utc": "2026-10-09T23:38:41.175848000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.