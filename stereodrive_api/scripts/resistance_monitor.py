"""Passive StereoDrive resistance telemetry monitor.

Reads USBPcap through TShark and writes timestamped AF12/AF45 replies. It never
opens the serial port and never sends controller commands. Run from an elevated
PowerShell because USBPcap capture normally requires Administrator rights.
"""
from __future__ import annotations

import argparse
import csv
import json
import subprocess
import time
from pathlib import Path

TSHARK = r"C:\Program Files\Wireshark\tshark.exe"


def monitor(interface: str, output: Path, duration: float | None) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        TSHARK, "-l", "-i", interface,
        "-Y", "usb.bus_id == 2 && usb.device_address == 3 && "
        "usb.transfer_type == 0x03 && usb.data_len > 0",
        "-T", "fields", "-E", "separator=,", "-E", "quote=d",
        "-e", "frame.time_epoch", "-e", "usbcom.data.out_payload",
        "-e", "usbcom.data.in_payload",
    ]
    started = time.monotonic()
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["epoch", "utc_iso", "direction", "opcode", "payload_hex",
                         "af12_byte5", "af12_byte6", "af45_payload_hex"])
        process = subprocess.Popen(command, stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE, text=True,
                                   encoding="utf-8", errors="replace")
        try:
            assert process.stdout is not None
            for line in process.stdout:
                if duration is not None and time.monotonic() - started >= duration:
                    break
                fields = next(csv.reader([line]))
                if len(fields) != 3:
                    continue
                epoch, out_payload, in_payload = fields
                if out_payload:
                    writer.writerow([epoch, time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(float(epoch))),
                                     "OUT", out_payload[:2], out_payload, "", "", ""])
                if in_payload:
                    raw = bytes.fromhex(in_payload)
                    b5 = raw[5] if raw.startswith(b"\x12") and len(raw) > 6 else ""
                    b6 = raw[6] if raw.startswith(b"\x12") and len(raw) > 6 else ""
                    writer.writerow([epoch, time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(float(epoch))),
                                     "IN", raw[:1].hex(), in_payload, b5, b6,
                                     in_payload if raw.startswith(b"\x45") else ""])
                handle.flush()
        finally:
            process.terminate()
            process.wait(timeout=5)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--interface", default=r"\\.\USBPcap2")
    parser.add_argument("--output", type=Path, default=Path("resistance-monitor.csv"))
    parser.add_argument("--duration", type=float, default=None,
                        help="seconds; omit to run until Ctrl+C")
    args = parser.parse_args()
    monitor(args.interface, args.output, args.duration)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
