# usb-20261009/axis-series/ap-negative001/ap-out-back-01.pcap

Scope: bus 2, devices 3 and 7. Filtered capture; original hash retained.

## Commanded

[
  {
    "command_id": "local-usb-701ec4386a0b432885391668eefd2f38",
    "kind": "out_and_back",
    "axis": "AP",
    "direction": -1,
    "step_mm": 0.01
  }
]

## Discovered

- Selector 0x40 absolute raw target 105854; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105907; profile words [1959, 8000, 8000]; tail 00010104.
- Selector 0x40 absolute raw target 105332; profile words [1959, 8000, 8000]; tail 00010104.
- Historical duplicate-click incident: two outward targets and one reversal; native return faulted. Not a successful return.

## Final observed counts

{
  "0x40": {
    "first_raw": 105802,
    "last_raw": 105332,
    "last_moving": 0,
    "samples": 28,
    "last_utc": "2026-10-09T21:57:58.370997000Z"
  },
  "0x50": {
    "first_raw": 75603,
    "last_raw": 75603,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T21:57:58.431855000Z"
  },
  "0x60": {
    "first_raw": 41819,
    "last_raw": 41819,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T21:57:58.492130000Z"
  },
  "0x70": {
    "first_raw": 322729,
    "last_raw": 322729,
    "last_moving": 0,
    "samples": 21,
    "last_utc": "2026-10-09T21:57:58.552663000Z"
  }
}

Status counts are not independent physical position feedback. See metadata.json for UTC, hashes, packets and limitations.