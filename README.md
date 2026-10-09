# Neurostar StereoDrive Craniotomy Planner

This repository contains a Windows Qt application for driving **StereoDrive** and the **Injectomate** during craniotomy drilling and injection workflows.

The main GUI is:

- [tools/craniotomy_qt.py](tools/craniotomy_qt.py)

The main Win32 automation/controller layer is:

- [tools/stereodrive_controller.py](tools/stereodrive_controller.py)

PowerShell diagnostics and direct control tests live in:

- [tools/control_stereodrive.ps1](tools/control_stereodrive.ps1)

## What The Software Does

The application sits on top of the native StereoDrive application and automates:

- AP / ML / DV movement
- Bregma reset / synchronization setup
- Seed generation for circular craniotomies
- Repeated drilling rounds around a perimeter
- Freeze / unfreeze marking of perimeter points
- Injection-site storage
- Automated injection sequences
- Injectomate syringe stepping and plunger-position tracking
- Injectomate calibrate-popup reading for real syringe position
- Benchmarking of axis movement speed
- Bregma-relative planning with an optional physical anchor reference
- Calibrated skull-overlay maps for craniotomy and injection planning
- Injection-grid generation and per-site validation
- Recoverable project-session autosave

It does **not** replace StereoDrive. StereoDrive must already be running.

## Requirements

- Windows
- Python 3
- `PySide6`
- StereoDrive running and detectable by its main window class/title
- Hardware connected and functioning in StereoDrive

## How To Run

From the repo root:

```powershell
python .\tools\craniotomy_qt.py
```

If StereoDrive is not running, the app will fail to attach.

## High-Level UI Layout

The window title is **Craniotomy Planner**.

Main areas:

- Top header for coordinate actions and quick stored locations
- **Setup** panel for craniotomy planning
- Two projection views for trajectory / seed visualization
- **Manual Control** panel for syringe manual actions and syringe-position display
- **Injection** panel for automated injection protocol settings and sequence progress
- **Injection Sites** panel for storing and resuming site lists
- **Options** dialog for keyboard bindings, StereoDrive control scans, USB probing, and update tools

## Startup Behavior

On startup the app:

- Connects to StereoDrive
- Starts a live position refresh timer
- Starts a background watcher that auto-confirms some StereoDrive dialogs:
  - below-skull warning: clicks **Yes**
  - "no actual movement to execute": clicks **OK**
- Schedules an early syringe-position update from the Injectomate calibrate popup
- Prompts to restore a recoverable project session when one exists. Active movement, drilling, and injection are **not** resumed automatically.

## Coordinate Controls

Top-row actions include:

- coordinate-mode buttons: `Axis` and `Bregma`
- `Set Bregma`, `Set Anchor`, and `At Anchor`
- `Bregma`
- `Home`
- `Work`
- `Go to`
- `Stop`
- `Drill On/Off` and `Clear Project`
- quick stored locations `A`, `B`, `C` and corresponding `Set A`, `Set B`, `Set C`

### Set Bregma

`Set Bregma` records the current **mechanical Axis** position as this GUI's local Bregma origin. It does **not** reset or modify StereoDrive's native Bregma reference. It clears any existing GUI anchor and selects Bregma mode.

The Axis/Bregma buttons select the live coordinate display and the coordinate frame of map clicks. Craniotomy plans, injection sites, and A/B/C positions are always stored in **GUI Bregma coordinates**, including when captured while displaying Axis coordinates. Set GUI Bregma before creating these project positions. The maps convert their stored points to Axis coordinates when Axis mode is selected.

Injection and drilling runs convert a copy of their targets to mechanical Axis coordinates using the current GUI Bregma origin before starting. This includes controlled descent, overshoot, continuous tracing, frozen-section retraction, and pause retraction. All arrival checks read StereoDrive's mechanical Axis fields. StereoDrive's own Bregma reference is not used for these movements.

### Anchor workflow

Use an anchor when Bregma is no longer physically accessible:

1. Set GUI Bregma at the reference point.
2. Move to a reproducible alternative point and press `Set Anchor`.
3. After a tool change, return to that physical anchor and press `At Anchor`.
4. Confirm the warning to recalibrate the GUI Bregma origin from the anchor offset.

`At Anchor` changes only the GUI's Bregma transform. It does not change the stored craniotomy or injection-site coordinates. The maps show the anchor with a blue anchor symbol; resetting GUI Bregma clears it.

### Bregma / Home / Work / Go to

- `Bregma` moves to the Bregma location
- `Home` moves to this GUI's saved mechanical Axis Home position
- `Work` moves to this GUI's saved mechanical Axis Work position
- `Go to` opens a coordinate dialog populated from the live position. If Bregma has been set, saved locations and manual entries in this dialog use Bregma coordinates even if the main display is currently in Axis mode.

Move to the desired Home or Work position, then open **Options** and press **Set Home** or **Set Work**. These positions are stored permanently in `settings.json` as mechanical Axis AP/ML/DV, remain unchanged by Bregma/anchor recalibration, and survive **Clear Project**. There is no fallback to StereoDrive's native Home/Work positions; unset positions must be saved first.

Every user-initiated navigation action shows a cancellable moving dialog with live map-cross updates. `Esc` or **Cancel Movement** requests a stop. Home/Work always send their saved Axis coordinates, even in Bregma display mode, and verify arrival from mechanical Axis readouts.

The top **Stop** cancels both drilling and injection workers and stops the manipulator and any running injection. Injection **Stop** also stops manipulator travel. Cancellation remains latched through target entry, GoTo delays, and retries, so a cancelled worker cannot issue a later GoTo. Moving dialogs confirm stable Axis readouts after stopping. Closing during an operation requests Stop and waits for its worker to finish; if it cannot confirm stopping, the window stays open.

One motor operation runs at a time. New navigation, benchmarks, probes, and reference changes are blocked during an active procedure. During the stationary site-validation dialog, normal movement and movement-step shortcuts remain available. Editable text/shortcut fields do not trigger motor shortcuts.

An opposite-direction movement key on an actively moving axis cancels that movement (and its running procedure), without executing the requested reversal. For example, press Left while ML is moving right, or Page Up while DV is moving down. Release and press again after stopping to move in the new direction; held-key repeats cannot restart the cancelled move. Requests in the same direction are not queued while the nudge is still in progress. Other new motor nudges remain blocked during a procedure.

### Saved Go to positions

The `Go to` dialog can save a named Bregma AP/ML/DV position, list previously saved positions, load one into the entry fields, or delete it. Saved locations are part of the recoverable project session.

### Quick Stored Locations

You can store three extra locations:

- `Set A` / `A`
- `Set B` / `B`
- `Set C` / `C`

These are intended for frequently revisited points and are stored in GUI Bregma coordinates in the current project session. After a tool change and **At Anchor**, they use the recalibrated GUI reference, regardless of the selected display mode.

### Clear Project

`Clear Project` removes the active craniotomy, injection sites, GUI Bregma/anchor calibration, and stored locations after confirmation. It preserves keyboard preferences and reusable craniotomy/injection settings.

## Keyboard Controls

When focus is not in an editable text field, the following shortcuts nudge the
current position:

| Key | Action |
|---|---|
| Left / Right | ML left / right |
| Up / Down | AP anterior / posterior |
| Page Up / Page Down | DV up / down |
| `Ç` | Increase movement step size |
| `Shift` | Decrease movement step size |
| `F1` / `F2` | Decrease / increase injection volume |
| `F3` / `F4` | Syringe step up / down |
| `Esc` | Stop injection |

The movement and syringe shortcuts can be reassigned in the **Options** tab.
Assignments are saved automatically in
`Documents/Neurostar_Master/Configs/settings.json` and restored when the app
restarts. Shortcuts are ignored while a text box or combo box is being edited.

## Maps, Overlay, And Map Moves

Both the Craniotomy and Injection tabs provide interactive maps. Select a calibrated skull image from the top-bar **Overlay** menu. The overlay, red current-position cross, and blue anchor symbol use Bregma coordinates and are therefore shown only in **Bregma** mode. Axis mode deliberately hides the overlay and displays a reminder to switch to Bregma mode.

Use the mouse wheel to zoom and left-button drag to pan. Scrolling out always reaches the full skull without going farther out. The craniotomy map provides **Zoom to craniotomy**, **Zoom to mid-range**, and **Zoom to skull**; the injection map provides craniotomy, injection-map, and skull focus modes.

Double-clicking a map point requests three stages: retract vertically at the current AP/ML, travel laterally at clearance height, then move vertically to the selected target. Bregma-mode map moves use clearance DV `-0.5 mm` and target DV `0 mm`; Axis-mode map moves use clearance 0.5 mm above the starting DV and return to the starting DV. If the tool is already higher, lateral travel uses that higher position. One cancellable moving dialog remains visible throughout and shows the current Axis target and stage. Navigation, injection validation, and protocol approaches also confirm vertical clearance before lateral travel. Missing DV targets are retried through the shared Axis arrival check in all procedures, unless live DV has already arrived.

## Project Recovery

Active projects are autosaved about once per second to `Documents/Neurostar_Master/Configs/project_session.json`. Recovery includes the GUI Bregma/anchor transform, craniotomy setup and captured surfaces, trajectory/drilling state, injection settings/sites, stored locations, selected overlay, map zoom choices, and active tab. The startup prompt can restore this project, but never resumes physical movement, drilling, or injection automatically.

New sessions record the coordinate frame explicitly. Older sessions without this metadata retain their saved data, but ambiguous craniotomies/sites cannot drive movement: recreate the craniotomy, clear/load injection sites, and re-save A/B/C positions. Named Go-to positions were already Bregma-defined and remain available. This avoids guessing which display mode was active when an old position was captured.

## Options And Updates

Movement regression checks can be run from the repository root with `python -B -m unittest discover -s tests -v`. They use simulated Windows controls and the GUI's original movement methods, with no hardware access. They cover references/tool changes, protocol targets, clearance, cancellation, target recovery, and session metadata. These checks do not replace testing the Windows GUI against the physical controller.

Open **Options** from the top-right button to:

- reassign or reset movement and syringe keyboard shortcuts;
- scan StereoDrive controls and copy the resulting control-ID report for diagnosis;
- run the USB Controller Probe described below; and
- run benchmark/diagnostic utilities.

The top-bar **Update** button pulls the latest GitHub version while discarding local repository changes. When the update succeeds, it offers to restart the application. Do not use it to preserve uncommitted source-code edits.

## USB Controller Probe

For an agent-controlled, supervised bench workflow on the same Windows
computer, use the separate [Local Movement Probe](tools/movement_probe/README.md).
It supports bounded Axis nudges, relative/absolute GoTo, fine/planar movement,
verified out-and-back experiments, and bounded injector injection/step/return
probes with a loopback-only HTTP API, no tokens,
arming or heartbeats, a local Stop button and UTC logs. Agent safety instructions are in
[its AGENTS.md](tools/movement_probe/AGENTS.md). It does not capture or replay USB.
The probe README also includes a
[two-paragraph agent handoff](tools/movement_probe/README.md#two-paragraph-agent-handoff)
with the current launch commands, token-free local API, Stop behaviour and
USB capture-analysis workflow.

The **Options → USB Controller Probe** is a manual correlation tool for
investigating the USB traffic produced by one StereoDrive axis nudge. It does
not listen to USB by itself, and it never replays or injects captured packets.
Instead, after an explicit confirmation it performs exactly one small positive
Axis nudge and, only after the Axis field confirms the expected movement,
performs the matching negative nudge. Its copyable log records local timestamps
and Axis readings for matching against the capture.

### Install packet capture on the StereoDrive Windows computer

1. Install [Wireshark](https://www.wireshark.org/download.html) using its
   official Windows installer. Select the optional **USBPcap** component if the
   installer offers it. If it does not, install the signed Windows installer
   from the [USBPcap project](https://desowin.org/usbpcap/). Administrator
   rights are required; restart Windows if the installer requests it.
2. Start Wireshark as appropriate for your local installation and select the
   USBPcap interface for the root hub that contains the StereoDrive controller.
   If the correct root hub is unknown, capture one at a time and identify the
   controller from its USB address and descriptor traffic.
3. Start the capture **before** running a single probe in Options. Run only one
   selected Axis/step pair at a time, then stop and save the capture.
4. Copy the probe log from the options dialog. Match the `FORWARD_NUDGE` and
   `REVERSE_NUDGE` timestamps with the adjacent USB request blocks in Wireshark.
   Record the controller USB address and VID/PID before applying display
   filters to later captures.

USBPcap records Windows USB request blocks (URBs), not electrical signals on a
USB wire. Use a dedicated hardware USB analyser if firmware-level or physical
line timing is required. Do not attempt to replay captured traffic to a live
stereotaxic controller: the diagnostic is intended only to observe a movement
that StereoDrive itself issued.

## Setup Panel

The **Setup** panel contains the craniotomy planning controls.

Important fields and actions include:

- craniotomy diameter
- seed count
- skull thickness
- max depth
- depth per round
- round time
- drill rate
- auto-start next round
- `Generate Seeds`
- `Next seed`
- `Set Surface`
- `Stop Motion`
- `Clear Surface Measurements`
- `Clear Craniotomy`
- `Start Drilling`
- freeze/unfreeze controls

### Generate Seeds

This creates a circular perimeter of drilling targets.

Behavior:

- the perimeter is drawn immediately
- the craniotomy circle is frozen visually once seeds are generated
- generating seeds resets current drilling progress state
- the app can ask whether to move to the first seed

### Set Surface

At each seed, `Set Surface` stores the surface DV for that point.

Special shortcut:

- `Ctrl + click Set Surface` sets all surface values to `0`

This is mainly for fast debugging.

### Clear Craniotomy

`Clear Craniotomy` removes the active seed plan, captured surfaces, trajectory, drilling progress, and frozen points. It leaves injection sites, Bregma/anchor calibration, and reusable setup values unchanged.

### Seed Navigation

`Next seed` moves through the seed list.

## Craniotomy Drilling Workflow

The drilling workflow is round-based.

For a round:

1. A current target depth is chosen.
2. The app traverses the circular perimeter.
3. At each point it moves to surface + current round depth.
4. Frozen points are skipped.
5. At the end of the round it returns above the center.

Frozen perimeter segments are shown in blue with a line three times thicker than the normal perimeter line.

After a round:

- the app beeps
- a countdown popup appears
- it can auto-start the next round after the countdown
- the next target depth can be edited in that popup
- if auto-start is enabled, it proceeds automatically
- otherwise it pauses

### Start / Pause / Continue Drilling

`Start Drilling` becomes `Pause` while a round is running.

Starting or manually continuing a drilling sequence asks **Is the drill turned on?** Choose **Yes** to proceed; **No** (the default), closing the prompt, or pressing Esc leaves drilling unstarted. This confirmation does not turn the drill on automatically. Subsequent automatically started rounds continue without repeating the prompt.

If paused:

- the drill is raised to `2 mm` above surface
- the button changes to `Continue`
- continuing returns to the first not-yet-drilled perimeter location for the current round depth

### Current Target Depth

There is a separate current target depth control under the drilling progress area.

- `Change` opens a dialog to set it directly
- this can be set deeper than skull thickness if needed

## Craniotomy Path Planning And Timing

Perimeter tracing uses a continuous AP/ML nudge path rather than simple point-to-point stepping.

Current behavior:

- AP/ML path stepping uses a DDA-style controller in `stereodrive_controller.py`
- the GUI estimates a `path_step_mm` for the round based on expected travel and round-time budget
- larger required travel or shorter round time results in larger step size and looser tolerance
- it prioritizes approximate path timing while still following the circle

Important limitation:

- `Round Time` is read when a round starts
- changing the round time during an already-running round does **not** currently re-plan that round
- the new value applies to the next round

## Manual Control Panel

The **Manual Control** panel contains syringe-related controls and readouts.

Main controls:

- `Step Syringe Up (F3)`
- `Step Syringe Down (F4)`
- red `Stop`
- `Empty Syringe`
- `Update Syringe Position`
- `Test for Blockage`

### Syringe Position

The vertical syringe gauge on the right tracks syringe position in `nl`.

Behavior:

- the full displayed range is `0` to `5000 nl`
- internal syringe state is tracked during manual stepping and injection programs
- on startup and before injection sequences, the real position is refreshed using the Injectomate calibrate popup

### Update Syringe Position

`Update Syringe Position` reads the Injectomate calibrate popup and updates the GUI/internal syringe position.

This uses the precise value exposed in the calibrate popup rather than trying to infer the on-panel custom gauge.

### Empty Syringe

This sends the Injectomate plunger to `0`.

## Injection Panel

The **Injection** panel defines a full single-site injection protocol.

Current settings:

- `Main injection volume (nl)`
- `Injection rate (nl/min)`
- `Insertion injection rate (nl/min)`
- `Injection depth (mm)`
- `Insert/retract speed (um/sec)`
- `Overshoot (mm)`
- `Post inject pause (s)`
- `Test volume (nl)`

Below that:

- program-sequence list, including total syringe volume required and expected timed duration
- overall sequence progress
- current injection / movement progress
- `Go`
- `Pause`
- `Stop`

### Injection Sequence Definition

For each site, the current protocol is:

1. Move to `1 mm` above the stored surface
2. Move normally to the stored surface
3. Insert from surface to `depth + overshoot` at `Insert/retract speed`
4. Inject during insertion at `Insertion injection rate`
5. Retract overshoot back to target depth at the same insert/retract speed
6. Continue main injection at target depth if needed
7. Post-injection pause at target if configured
8. Retract to the stored surface at insert/retract speed
9. Move normally to `1 mm` above the stored surface
10. Optionally run a blockage test

### Important Injection Details

- insertion overshoot and final injection depth are explicitly reached even when timed sampling or syringe calls span a movement phase boundary
- insertion and retraction are handled by controlled DV motion
- the pipette advances downward while injection is occurring
- overshoot defaults to `0.05 mm`
- the post-injection pause is shown in the feedback/status area
- the active sequence step is bolded in the program-sequence list during execution
- the active injection site is bolded in the injection-site list during execution
- total syringe planning includes main injection volume, insertion volume, and two test volumes per site when blockage checks are enabled

## Injection Sites Panel

This panel manages stored injection sites.

Controls:

- `Add Injection Site`
- `Add Grid`
- `Save Site Set` / `Load Site Set`
- `Nudge All Sites`
- `Remove Selected Site`
- `Validate Sites`
- `Clear Sites`
- `Start From Selected`
- `Check blockage after each site`

### Injection Site Storage

Each stored site contains:

- AP
- ML
- surface DV

If no sites are stored, the current location is treated as the active site and its current DV is treated as the surface.

Site capture always converts mechanical Axis readings to GUI Bregma coordinates. Injection execution uses a frozen copy converted back to Axis coordinates with the latest calibration, including after **At Anchor**; selecting Axis display mode does not change the injection destinations.

### Grid sites and validation

To place individual sites visually, enable **Add sites on map** below the Injection map, then single-click the desired locations. Each click adds a gray, unvalidated site with AP/ML in GUI Bregma coordinates and no surface DV. GUI Bregma must be set first; clicks in Axis display mode are converted back to GUI Bregma coordinates. Sites appear immediately on the map/list and are autosaved. Use **Validate Sites** to fine-tune them and capture their surfaces before injection.

In this mode, clicks add sites without moving the tool, dragging still pans, and the mouse wheel still zooms. Double-clicks do not request movement or add a second site. Turn **Add sites on map** off to restore double-click-to-move behavior. Placement mode starts off when the app opens and is cleared by **Clear Project**.

`Add Grid` creates a Bregma-centred AP/ML grid. The dialog accepts AP-site count, ML-site count, AP spacing, ML spacing, and recently used configurations. Grid sites intentionally have no surface DV and appear light gray until validated.

`Save Site Set` saves Bregma AP/ML site targets. `Load Site Set` replaces the current list with the saved targets, deliberately marking all of them unvalidated so surface location can be rechecked for the current preparation.

Set **Validation height (mm above Bregma)** in the Injection Sites controls to choose the approach height for all validations. It defaults to `0.5 mm` (GUI Bregma DV `-0.5 mm`); increase it when the surface bulges above the reference. The value is saved permanently in app settings.

Right-click any site in the list and choose **Validate this site** to approach just that site and open the normal adjustment dialog. **Validate**, **Skip**, **Delete Point**, or **Cancel** ends this single-site pass without automatically moving to another site. The active site is highlighted blue during validation. This works in either coordinate display mode; stored coordinates remain GUI-Bregma-relative.

Choose `Validate Sites` then either **Validate all** or **Validate unvalidated**. For each site the app retracts vertically before travelling to AP/ML at the configured height, opens a modal validation dialog, and keeps the normal movement/speed keyboard shortcuts active. Adjust AP/ML/DV to the desired surface location, then choose:

- **Validate and Next**: stores the refined AP/ML/surface DV and proceeds;
- **Next Without Validating**: keeps the site unchanged/unvalidated;
- **Delete Point**: removes it; or
- **Cancel**: stops the validation pass.

Injection cannot start while any stored site is unvalidated.

When `Nudge All Sites` is enabled, the AP/ML keyboard movement shortcuts translate every listed site by the current movement-step size without moving the manipulator. This marks all sites unvalidated, requiring a new validation pass.

## Starting An Injection Sequence

Press `Go` in the Injection panel.

Before starting, the app:

- reads current GUI injection settings
- syncs real syringe position from the calibrate popup
- estimates required syringe volume including optional blockage tests
- refuses to start if the requested sequence would exceed syringe limits

## Stopping And Resuming Injection

### Stop

Pressing `Stop`:

- requests sequence stop
- attempts to stop active Injectomate motion
- refreshes real syringe position from the calibrate-popup read method

### Resume From Selected

After a stopped sequence:

1. Select a site in the `Injection Sites` list
2. Click `Resume From Selected`

The app will:

- restart the protocol from the selected site onward
- use the **current GUI settings**
- re-check syringe position before restarting

This is not a low-level resume inside a partially executed site. It resumes at the selected stored site boundary.

## Blockage Test Behavior

If `Check blockage after each site` is enabled:

- after returning to `1 mm` above the stored surface, the app performs a test injection
- the blockage prompt uses a gentle repeating alert sound to draw attention
- once the test injection completes, it asks:
  - `Is the test injection confirmed not blocked?`

If you choose:

- `Not blocked`: continue to the next site
- `Blocked`: the app asks whether to do another test injection
- `Another test`: it repeats the blockage test at the same site
- `No`: it stops the sequence there

This repeats until:

- you confirm clear, or
- you decline further tests, or
- you stop the sequence manually

## Syringe Limits And Safety Checks

The GUI tracks syringe position in `nl` and enforces bounds.

Current allowed range:

- minimum: `0 nl`
- maximum: `5000 nl`

If a requested syringe movement or full sequence would exceed those bounds:

- the app shows a warning
- the action is rejected
- for sequence planning, the required total volume is checked in advance

## Automatic StereoDrive Dialog Handling

The app runs a background watcher that auto-confirms common dialogs:

- below-skull warning: clicks `Yes`
- no-actual-movement dialog: clicks `OK`

This helps keep automated drilling and movement from hanging on routine dialogs.

## Benchmarking

The `Benchmark` button runs a movement benchmark and shows the results in a popup text box suitable for copy/paste.

This is intended to measure actual movement speed for:

- AP
- ML
- DV

across several step sizes.

The benchmark is useful for tuning round-time expectations and continuous path behavior.

## Notes On Injectomate Control

The code can:

- show the Injectomate panel directly
- read the calibrate-popup plunger value
- trigger syringe steps up/down
- stop active Injectomate motion

The most reliable precise syringe-position read path is currently the **Injectomate calibrate popup**.

## Troubleshooting

### The app says StereoDrive was not found

- Start StereoDrive first
- Make sure the StereoDrive main window is actually present

### The app opens but movement does not happen

- Confirm StereoDrive is connected to the hardware
- Confirm the selected nudge sizes are valid in StereoDrive
- Watch for any external modal dialogs not covered by the auto-confirm watcher
- For a map move, read the Axis AP/ML/DV target shown in the moving dialog. The app verifies arrival from StereoDrive's live Axis values rather than relying solely on the target text boxes.

### Set Bregma does not work

- `Set Bregma` requires readable live Axis values from StereoDrive.
- It does not modify StereoDrive's native Bregma reference.
- If the live Axis fields cannot be read, bring the main StereoDrive window to the foreground and retry.

### Injectomate position looks wrong

- Press `Update Syringe Position`
- This re-reads the calibrate popup and overwrites the internal estimate with the real plunger value

### The sequence stops after a blockage test

- This is expected if you answered that the pipette is blocked and declined another test injection
- Select the next desired site and use `Resume From Selected` if needed

## Main Files

- GUI: [tools/craniotomy_qt.py](tools/craniotomy_qt.py)
- Controller: [tools/stereodrive_controller.py](tools/stereodrive_controller.py)
- PowerShell diagnostics: [tools/control_stereodrive.ps1](tools/control_stereodrive.ps1)

## Typical Usage

A common session looks like this:

1. Start StereoDrive
2. Run `python .\tools\craniotomy_qt.py`
3. Press `Set Bregma` at the intended local reference point
4. Set craniotomy diameter and seed count
5. Press `Generate Seeds`
6. Move around the perimeter and capture surface values with `Set Surface`
7. Start drilling rounds with `Start Drilling`
8. Add/load injection sites, then validate each site before injection
9. Configure the injection protocol
10. Press `Go`
11. Confirm blockage-test outcomes between sites if enabled

## Current Scope

This README describes the current behavior implemented in the codebase as of the present repository state. It is an operator guide, not a formal validation document.
