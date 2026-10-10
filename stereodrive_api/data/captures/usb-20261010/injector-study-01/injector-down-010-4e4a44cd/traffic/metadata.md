# usb-20261010/injector-study-01/injector-down-010-4e4a44cd/traffic.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "injector-usb-7fd2eaa93ca54a4cb7fef4e86887062f",
    "kind": "injector_step",
    "direction": "down",
    "volume_nl": 10
  }
]

## Discovered

- Selector 0x70 absolute raw target -8572; profile words [1820, 8000, 8000]; tail 00010004.
- Native completion preceded final USB idle; capture ends moving. Later separate passive capture verified original count -8572 idle. Do not claim final offset.

## Final observed counts

{
  "0x40": {
    "first_raw": 103085,
    "last_raw": 103085,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T23:30:18.329256000Z"
  },
  "0x50": {
    "first_raw": 78320,
    "last_raw": 78320,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T23:30:18.388828000Z"
  },
  "0x60": {
    "first_raw": 44379,
    "last_raw": 44379,
    "last_moving": 0,
    "samples": 4,
    "last_utc": "2026-10-09T23:30:18.449437000Z"
  },
  "0x70": {
    "first_raw": -1133,
    "last_raw": -8324,
    "last_moving": 1,
    "samples": 9,
    "last_utc": "2026-10-09T23:30:18.548162000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.