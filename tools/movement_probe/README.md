# Same-computer direct-API movement probe

This branch uses the public USB API, not native StereoDrive controls. Close
StereoDrive, the planner and all other controllers before hardware connection.
The server accepts same-computer requests immediately once verified setup connects:
no token, arm or heartbeat. Bind is restricted to `127.0.0.1`; never tunnel/proxy it.
Read [AGENTS.md](AGENTS.md) before hardware use.

## Run

Python 3.10+ with Tkinter (standard Windows Python installer), no pip dependencies.
From the repository root, simple fake-controller smoke test:

```powershell
py -3 tools\movement_probe\server.py --simulate
```

For actual API simulation, provide `--simulate --api-setup PATH`. Hardware:

```powershell
py -3 tools\movement_probe\server.py --api-setup C:\verified\probe-setup.json
```

Optional `--allow-dv` and `--allow-injector` require independent clearance/syringe
verification. Drill ON is not provided. `--console` uses a terminal instead of GUI;
use `pyw -3` for the GUI without terminal after verifying startup. Stop button/Esc,
client Stop or window close request all channel Stops and drill OFF.

The local setup JSON contains:

- `calibration`: measured `axis_zero_counts` (AP/ML/DV), `piston_3000_count` and
  `anchor_backlash` (AP/ML/DV/PISTON), as documented in the [API](../../stereodrive_api/README.md).
- `verified_backlash`: all four independently verified current states, separate
  from calibration anchor states. Unknown history must not be guessed.
- Optional `speed_mm_s`: 1 (default) or 2 only.
- Optional `new_reference`: false (default). True only after independent reference
  verification, not to bypass an unresolved fault or enlarge the travel envelope.

No live sample counts are supplied: obtain measured calibration first. Direct state
is outside the repository in `%LOCALAPPDATA%\NeurostarDirectProbe`, with separate
live/simulation files. Connection refuses mismatched/invalid state.

## Commands and limits

```powershell
py -3 tools\movement_probe\client.py status
py -3 tools\movement_probe\client.py move --json-file tools\movement_probe\example_ap_out_and_back.json
py -3 tools\movement_probe\client.py events
py -3 tools\movement_probe\client.py stop
```

HTTP: GET `/status`, GET `/events`, POST `/move`, POST `/stop`. POST uses JSON and
`Content-Type: application/json`. Movement requires unique `command_id`. Poll the
returned operation; for an uncertain response retry only identical ID/body in the
same session, never a new ID. Axis kinds: `nudge`, `axis`, `relative`, `absolute`,
`out_and_back`; methods `goto`, `nudged`, `planar`. Coordinates are calibrated
mechanical Axis mm, not planner/native Bregma. GoTo is sequential, not clearance planning.

Axis bounds are ±1 mm per axis from connection/startup; max 1 mm Euclidean distance
per command/leg. DV is disabled unless opted in. Begin with 0.01 mm AP. Completion
requires API exact raw target/idle settlement and server stable arrival, not a
single rounded position. Timeout is 60 seconds for the complete experiment.

With `--allow-injector`, supported kinds are `injector_step` and
`injector_out_and_back`, using `direction: "up"/"down"`, `volume_nl` and `command_id`.
Volumes: 10/20/50/100 nL, hard maximum 100 nL per leg; estimated piston ±100 nL
from connection and 0–5000 nL capacity. `injector_inject` is refused by the direct
backend: controlled-flow injection is unvalidated. The simple fake smoke simulator
retains that historical command only for HTTP tests; it does not prove API support.
Free-piston completion is not measured delivery; reversing may aspirate air and
cannot undo injection. Never request fill/empty or rate/type changes.

The GUI shows incoming requests, motor-derived positions, outcomes and errors.
UTC JSONL logs: `%USERPROFILE%\StereoDriveProbeLogs`. USB capture is separate;
read only operator-approved saved captures with TShark `-r`. No guessed packet replay.

## Two-paragraph agent handoff

This branch has changed the movement server from native GUI automation to the
shared direct USB API. Close StereoDrive and every other controller first. The
operator must supply measured zero calibration and independently verified current
backlash history in `--api-setup`; DV/piston require explicit opt-ins. Test
`--simulate` first, then actual API simulation with the setup file. Once running,
the server accepts requests at `http://127.0.0.1:8765` without tokens/arming/heartbeat.
Use the client commands above or the four HTTP endpoints; poll each operation and
retry ambiguous responses only with identical ID/body in the same session.

Operate only with human-authorized supervised bench clearance, no specimen, drill
off and physical Stop accessible. Start at AP 0.01 mm, one command at a time;
limits are ±1 mm from startup and 1 mm per leg. Coordinates are calibrated Axis,
not Bregma. Piston free steps are 10–100 nL after syringe/plunger verification;
controlled injection is disabled. Software Stop is best effort, not an emergency
stop; client loss does not trigger heartbeat cancellation. Faults/unknown history
require independent verification, not new-reference/reconnect tricks. Preserve
logs/captures, obtain physical observations and ask the operator on uncertainty.
