# Supervised network movement probe

A small, separate console application for an agent to request controlled
StereoDrive movements while the operator records USB traffic. It reuses our
Windows GUI-control adapter; it **does not capture USB, replace the USB driver,
or replay/inject packets**. No PySide6 or third-party Python packages are needed.
Use Python 3.10 or newer. Real operation requires Windows, StereoDrive running
in the same interactive desktop session, and matching privilege levels (if
StereoDrive is elevated, the server must be too).

Read [AGENTS.md](AGENTS.md) before allowing an agent to run experiments.
This is a supervised bench diagnostic, **not an autonomous surgical controller**.
Do not use it with a specimen or a tool in contact with anything.

## Movement and limits

All values are **mechanical Axis millimetres**, never StereoDrive Bregma or the
main GUI's Bregma. AP/ML/DV map to the existing controller's axis directions;
positive/negative are numerical directions, not anatomical guesses.

Available movement commands:

| kind | fields | behaviour |
| --- | --- | --- |
| `nudge` | `axis`, `direction` (+1/-1), `step_mm` | One native axis arrow click; step 0.01, 0.02 or 0.05 mm |
| `out_and_back` | same as nudge | Verify outward arrival, then request and verify reverse; no blind reverse after failure |
| `axis` | `axis`, `target_mm`, optional `method` | Single-axis absolute target |
| `relative` | `coordinates` object, optional `method` | Signed offsets from the live mechanical readings |
| `absolute` | `coordinates` object, optional `method` | Mechanical Axis targets; omitted axes retain live values |

Methods for axis/relative/absolute:

- `goto` (default): native mechanical Axis text fields and GoTo, potentially
  multi-axis motion. Targets rounded to 0.01 mm, arrival checked within 0.006 mm.
- `nudged`: AP then ML then DV using verified 0.01 mm increments, not continuous
  drilling. Endpoint precision is approximately 0.01 mm.
- `planar`: AP/ML interleaved verified 0.01 mm increments, selecting the axis
  with the largest remaining distance. DV changes prohibited. Not an exact
  reproduction of the main GUI's continuous DDA drilling trajectory.

Default startup envelope: ±0.1 mm on each axis around startup position.
Per-command Euclidean distance: at most 0.05 mm. Out-and-back applies this limit
to each leg. Limits cannot be changed over the API. Local CLI can set radius
up to 1 mm and per-command distance up to 0.1 mm; increase only after a human
has physically verified the entire space. DV changes require local `--allow-dv`.
No syringe, drill power, native Home/Work, Bregma reset, or calibration endpoints.
Bookmarks may be tested using absolute coordinates only if within the envelope.

There is **no automatic skull-clearance move**: it would contaminate the capture
and require unapproved travel. The operator must establish an empty, clear
workspace before startup. Collision safety cannot be inferred from coordinates.

## Windows setup and first test

From the repository root in PowerShell:

```powershell
py -3 -B tools\movement_probe\server.py --simulate
```

This uses fake positions and issues no hardware commands. Check this workflow
before real experiments. The console prints a fresh token and starts DISARMED.
In a second terminal, set the token locally (do not paste it into shared logs):

```powershell
$env:STEREODRIVE_PROBE_TOKEN = 'TOKEN_FROM_SERVER_CONSOLE'
py -3 -B tools\movement_probe\client.py status
py -3 -B tools\movement_probe\client.py heartbeat --seconds 60
```

While that heartbeat is running, the **human operator** types `ARM` in the
server console. Initial heartbeat replies while disarmed are expected. Once
armed, use another terminal with the same token environment:

```powershell
py -3 -B tools\movement_probe\client.py move --json-file tools\movement_probe\example_ap_out_and_back.json
py -3 -B tools\movement_probe\client.py status
py -3 -B tools\movement_probe\client.py events
py -3 -B tools\movement_probe\client.py stop
```

`move` returns immediately with the operation ID. Poll `status` until its
operation is `completed` or `stopped`; accepting the HTTP request does not mean
arrival. The heartbeat command expires after 60 seconds by default and requests
Stop on exit. Explicit `/stop`, heartbeat loss (>3 seconds), motion timeout
(15 seconds), safety/readout failure, or leaving the envelope disarms the app.
There is no remote arm endpoint. Rearming requires the local console.

Only run one experiment and one heartbeat client at a time. A heartbeat is
not permission to leave the apparatus unattended. Stop that process immediately
if agent supervision is lost; do not leave an orphan heartbeat process running.

### Real hardware

1. Close the main craniotomy/injection GUI and stop other automation. Leave
   StereoDrive visible. Human verifies motors idle, tool well clear, no specimen,
   drill off, and physical Stop accessible. Do not change tools during a session.
2. Install USBPcap/Wireshark separately if capturing, restart Windows as required,
   select the controller's USB root hub, and start a short capture before moving.
3. Start the server **without** `--simulate`:

   ```powershell
   py -3 -B tools\movement_probe\server.py
   ```

4. Check the startup Axis readings against StereoDrive and approve the bounds.
   Start heartbeat; human types `ARM`. Begin with AP 0.01 mm out-and-back only.
5. Check measured return in `/status` and visually. Stop, save the USB capture,
   and preserve the matching `.jsonl` log before testing another variable.
6. DV experiments need a separate locally approved launch with `--allow-dv`.
   Do not infer anatomical up/down from the sign; verify mechanically first.

The server **never accepts StereoDrive's below-skull confirmation**. If detected,
it stops and disarms; the operator must investigate. No automatic target retry
is performed when arrival is ambiguous, and no recovery movement is issued on
cancel/error. Native direct-entry behaviour can begin motion before GoTo; every
requested target is bounded regardless. GUI Stop is best-effort, not a hardware
emergency stop. Windows/UI hangs or power/network failures can defeat software
Stop; use the physical Stop immediately if movement persists.

## Local network access

Prefer loopback (`127.0.0.1`) when the agent is on the controller computer.
For a different computer on a trusted private LAN, bind **this Windows
computer's specific private IPv4**, for example:

```powershell
py -3 -B tools\movement_probe\server.py --bind 192.168.1.50
```

Client commands then add `--url http://192.168.1.50:8765`.
If blocked, have the operator/administrator allow TCP 8765 only on the Windows
Private firewall profile and only from the agent computer's IP. Do not open the
Public profile, forward a router port, disable the firewall, or expose this
service to the internet. The server rejects wildcard/public bind addresses.
HTTP tokens are **not encrypted**; use loopback or an encrypted tunnel if the
network is not trusted. Do not send credentials via URL/query strings.

## API for the agent

Every request requires `Authorization: Bearer TOKEN`. No browser-origin requests
are accepted. JSON POST bodies must be objects of at most 4096 bytes.

| endpoint | result |
| --- | --- |
| `GET /status` | Actual Axis position, origin/bounds, armed state, active/last operation, Stop error |
| `GET /events` | Timestamped experiment log |
| `POST /heartbeat` with `{}` | Keep an already locally armed session alive |
| `POST /move` | Request one movement; rejects busy, disarmed, invalid or out-of-bounds commands |
| `POST /stop` with `{}` | Request native Stop and disarm, even if already disarmed |

All movement bodies require a unique `command_id`. Retransmitting the exact
same body/ID returns that operation without re-executing it. Reusing an ID with
different contents is rejected. IDs are retained for this server session only;
after a restart never retry an uncertain command. Maximum 1000 operations per
session. HTTP 400 means rejected; 401 authentication; 403 browser origin;
503 controller/server failure. Read the JSON error, do not assume success.

Examples (send as JSON via `--json-file`; create a new ID for each experiment):

```json
{"command_id":"ap-negative-01","kind":"nudge","axis":"AP","direction":-1,"step_mm":0.01}
```

```json
{"command_id":"xy-relative-01","kind":"relative","coordinates":{"AP":0.01,"ML":0.01},"method":"planar"}
```

```json
{"command_id":"ap-fine-01","kind":"axis","axis":"AP","target_mm":31.14,"method":"nudged"}
```

Absolute values above are illustrative, **not instructions to move to them**.
Always derive actual targets from `/status` and approved bounds.

Logs are saved immediately per event to `%USERPROFILE%\StereoDriveProbeLogs`,
outside the repository, with UTC timestamps, sequence numbers, requests, targets,
measured arrivals, and Stop/error events. Startup token is not logged. USB traces
are saved separately in Wireshark. Match `MOVE_REQUEST`, `ARRIVED`, and
`REVERSE_REQUEST` timestamps to USB transfers. Keep an idle baseline and change
only one axis/distance/method per capture. Logs are observations, not proof of
the USB protocol's meaning. Do not replay a guessed command to test a hypothesis.

Tests from repo root:

```powershell
py -3 -B -m unittest discover -s tests -v
```

Simulation tests validate HTTP/control safeguards, not physical collision
safety, Windows driver compatibility, or real StereoDrive arrival timing.

## Automated reading of Wireshark captures

An agent on the Windows machine can read saved USB `.pcap`/`.pcapng` files with
Wireshark's `tshark.exe`; no clicking in Wireshark is needed. This is Wireshark,
not WireGuard (a VPN whose diagnostics do not describe USB traffic).
See the [official TShark manual](https://www.wireshark.org/docs/man-pages/tshark.html).

After the capture has stopped and its file has been saved, for example:

```powershell
& 'C:\Program Files\Wireshark\tshark.exe' -r 'C:\captures\ap-001.pcapng' -Y usb -T json -x
```

`-r` reads an existing file, `-Y usb` selects USB packets, and `-T json -x`
produces machine-readable decoded fields with packet hex. Have the agent parse
stdout, correlate packet timestamps against the probe's UTC JSONL events, and
report candidate command bytes versus idle polling and device replies. Start
with complete USB records rather than assuming a payload field/endpoint before
the device is identified. A zero-packet result is not evidence of no commands:
verify interface/device selection and decoding first.

The agent may automate read-only analysis of agreed capture files. Do not use
`-i` to start a new live capture unless the operator has authorized the capture
interface/scope. Never use packet interpretation as authority to inject/replay
commands. Preserve originals; write any derived reports to a separate folder
outside the repo, avoiding accidental capture of other USB devices' private data.
