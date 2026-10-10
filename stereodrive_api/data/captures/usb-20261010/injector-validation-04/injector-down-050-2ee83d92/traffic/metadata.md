# usb-20261010/injector-validation-04/injector-down-050-2ee83d92/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "piston-range-31eeab7d70c04429b744f03abfcbc6aa",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 50
  }
]

## Discovered

- Selector 0x70 absolute raw target -504; profile words [1820, 8000, 8000]; tail 00010004.
- Piston down 50 nL: delta -13893 counts; last two samples independently idle at endpoint before next leg.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:44.584780000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:44.615550000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 7,
    "last_utc": "2026-10-09T23:42:44.646341000Z"
  },
  "0x70": {
    "first_raw": 13389,
    "last_raw": -504,
    "last_moving": 0,
    "samples": 18,
    "last_utc": "2026-10-09T23:42:44.676737000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.