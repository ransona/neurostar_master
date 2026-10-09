# Same-computer movement probe

A small GUI for an agent on the StereoDrive Windows computer to request
controlled movements while recording USB traffic. **Starting the server enables
local requests immediately: no token, Arm button, or heartbeat.** It binds only
to `127.0.0.1`; LAN, public and wildcard addresses are rejected.

This reuses the Windows GUI adapter; it does not capture USB, replace drivers,
or replay/inject packets. Use Python 3.10+ with Tkinter (select **tcl/tk and IDLE**
in the Windows Python installer). No pip packages needed. Server and StereoDrive
must share an interactive desktop session and privilege level. Read AGENTS.md.
Supervised bench use only: no specimen, tool safely clear, drill off, physical
Stop accessible. Close the main craniotomy GUI and other automation first.

## Two-paragraph agent handoff

The movement probe now runs as a small GUI and accepts commands immediately
while running: no token, Arm step, or heartbeat. It only accepts connections
from the same Windows computer at `http://127.0.0.1:8765`, not from the LAN.
From the repository root, first run `py -3 tools\movement_probe\server.py --simulate`;
for approved real experiments, close the main controller GUI, leave StereoDrive
visible, and run without `--simulate`. Use `client.py status`,
`client.py move --json-file PATH`, `client.py events`, and `client.py stop`
with the full prefix `py -3 tools\movement_probe\`. HTTP equivalents are
GET /status, GET /events, POST /move, and POST /stop; POST bodies require JSON
and Content-Type application/json. Each movement needs a unique command_id;
poll that operation until completed/stopped. Retry an uncertain response only
with the identical ID/body within the same server session. The window shows
incoming commands, positions, results and errors; all coordinates are mechanical
Axis mm, never native or main-GUI Bregma.

Operate only within the human-approved, supervised bench setup: no specimen,
tool clear, drill off and physical Stop accessible. Begin with the supplied
AP 0.01 mm out-and-back example; default limits are ±1 mm per axis from startup
and 1 mm total distance per movement, with DV disabled unless the operator approves --allow-dv.
One command at a time; do not bypass limits, dismiss skull warnings or replay
USB packets. Stop via client.py stop, POST /stop, the GUI Stop button or Esc;
another explicit command is allowed after cancellation finishes, while closing
the server disables all requests. There is no heartbeat-based disconnect Stop:
an accepted movement may continue until arrival or its 15-second timeout.
Hardware faults require investigation and a local restart. Capture USB separately
with Wireshark/USBPcap and read agreed saved captures using TShark -r; correlate
them with UTC JSONL logs in %USERPROFILE%\StereoDriveProbeLogs. Preserve evidence
and ask the operator before recovering from an unexpected result.

## Run it

From PowerShell in the repo root, test without hardware:

```powershell
py -3 tools\movement_probe\server.py --simulate
```

For real hardware, leave StereoDrive visible:

```powershell
py -3 tools\movement_probe\server.py
```

Use `pyw -3` instead to avoid a terminal window after verifying startup.
`--console` enables terminal-only `STOP`, `STATUS`, `QUIT` (use `py`, not `pyw`).

The GUI shows READY, URL, live Axis readings/bounds, commands, sender IPs,
accepted/rejected results and errors. **Show status polling** includes routine
requests. Full UTC JSONL logs go to `%USERPROFILE%\StereoDriveProbeLogs` outside
the repository. No environment variables, token-copy or arming steps.

The agent on the **same computer** can immediately run:

```powershell
py -3 tools\movement_probe\client.py status
py -3 tools\movement_probe\client.py move --json-file tools\movement_probe\example_ap_out_and_back.json
py -3 tools\movement_probe\client.py status
py -3 tools\movement_probe\client.py events
py -3 tools\movement_probe\client.py stop
```

`move` returns an operation ID immediately. Poll status until that ID completes
or stops; acknowledgement is not arrival. The sample verifies AP +0.01 mm then
reverses and verifies return. No blind reversal if the outward move fails.

**STOP Movement**, **Esc** or POST /stop cancels the current move. Once its
worker exits, another explicit command is allowed without arming/restarting.
**Close the server to prevent all further requests.** Closing also sends Stop.
Failures (readout/controller, skull warning, leaving bounds, failed Stop or
15-second timeout) latch FAULT; resolve and restart locally. There is no remote
fault reset or limit change. Client disconnection does not cancel an accepted
move: it may finish or run until timeout. Use local/physical Stop if needed.

## Movements and bounds

Values are **mechanical Axis mm**, never native or main-GUI Bregma. Verify
numerical directions mechanically; signs do not prove anatomical up/down.

| kind | fields | behaviour |
| --- | --- | --- |
| nudge | axis, direction (+1/-1), step_mm | One arrow click: 0.01, 0.02, 0.05, 0.1, 0.2, 0.5 or 1 mm |
| out_and_back | same as nudge | Verified outward nudge and matching reverse |
| axis | axis, target_mm, optional method | Single-axis absolute target |
| relative | coordinates object, optional method | Signed offsets from live readings |
| absolute | coordinates object, optional method | Axis targets; omitted axes stay put |

Methods: `goto` (default) uses Axis text fields/GoTo and may move multiple axes;
0.01 mm target resolution, 0.006 mm arrival tolerance. `nudged` performs AP then
ML then DV via verified 0.01 mm increments. `planar` interleaves AP/ML increments
by greatest remaining distance; DV forbidden. It is not continuous drilling DDA.

Default bounds: ±1 mm per axis from startup. Maximum Euclidean distance per
command/each probe leg: 1 mm. These are also the hard maximums; use
`--radius-mm`/`--max-move-mm` to choose smaller limits (max move must not exceed
radius). For multi-axis requests the 1 mm limit is total distance, not 1 mm on
every axis simultaneously. Restarting establishes a new startup position.
DV requires `--allow-dv`
at locally approved startup. No drill, syringe, native Home/Work, Bregma reset
or calibration API. No automatic skull-clearance/recovery moves. Bounds cannot
prove collision safety. Direct-entry builds may move before GoTo; targets are
bounded either way. Skull warnings are never accepted; no ambiguous-arrival
retry. GUI Stop is best-effort, not a hardware emergency stop.

## HTTP API

URL: `http://127.0.0.1:8765` (`--port` changes port). No authentication needed.
Any process on this computer can request bounded moves while it runs: only run
trusted software. Browser-origin requests, non-loopback Host headers and
non-JSON POSTs are blocked. Never expose/proxy/forward/tunnel this service to
other computers or change its binding/firewall restrictions.

| endpoint | result |
| --- | --- |
| GET /status | Axis position, origin/bounds, ready/fault state and active/last operation |
| GET /events | Timestamped request/command/results log |
| POST /move | One bounded movement; rejects busy/invalid/faulted requests |
| POST /stop with {} | Native Stop; new explicit requests allowed unless faulted |

POST requires `Content-Type: application/json`, JSON object of 1–4096 bytes.
Movement requires unique `command_id`. Exact same ID/body returns the original
operation without repeating it; different body with same ID is rejected.
IDs last for the server session (max 1000); never retry uncertain commands after
a restart. Errors: 400 invalid/busy/faulted, 403 browser/nonlocal Host, 415
non-JSON, 503 controller failure. Read the error; do not assume success.

Example JSON (use a fresh ID for each new experiment):

```json
{"command_id":"xy-relative-01","kind":"relative","coordinates":{"AP":0.01,"ML":0.01},"method":"planar"}
```

Use --json-file to avoid Windows quoting issues. Derive absolute targets from
live status, never arbitrary examples.

## Automated Wireshark reading

Install USBPcap/Wireshark separately; choose the controller root hub and capture
before moving, then stop/save. Keep idle baseline and one changed variable per
capture. Match JSONL MOVE_REQUEST, ARRIVED, REVERSE_REQUEST UTC events to USB.
The agent can read agreed saved captures with TShark, not WireGuard VPN logs:

```powershell
& 'C:\Program Files\Wireshark\tshark.exe' -r 'C:\captures\ap-001.pcapng' -Y usb -T json -x
```

Parse decoded USB fields/hex and distinguish requests/replies/idle traffic before
hypothesizing commands. Zero packets may mean wrong interface/device or decoding.
See [TShark manual](https://www.wireshark.org/docs/man-pages/tshark.html).
Live capture needs separately agreed scope; analysis never authorizes replay.
Preserve originals, keep reports outside the repo and avoid exposing unrelated
USB devices' private traffic. Test with `py -3 -B -m unittest discover -s tests -v`.
Simulation tests control logic, not physical safety or real Windows timing.
