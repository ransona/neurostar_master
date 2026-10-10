# Validation completed 10 October 2026

- 21 hardware-free API and hidden-window GUI tests passed.
- Coverage includes captured packet profiles; absolute calibration and anchor backlash;
  restored position; state invalid before writes; stale/corrupt state refusal;
  no target retry after bad reply; cumulative limits; opt-ins; concurrent motion query;
  Escape cancellation; three same-direction piston moves; drill power state, ON opt-in,
  error OFF attempt and Stop OFF; GUI routes commands through the shared API.
- All 87 archived captures have JSON and readable Markdown metadata; capture SHA-256
  values match the manifest. Report/document links and all five rendered figures checked.
- GUI opened disconnected in Simulation and its controls were visually inspected.
- No hardware motion or drill ON was performed while developing/testing this update.

These checks establish software behavior, not missed-step safety, physical tool
accuracy, delivered fluid volume, RPM or full acceptance of direct drill control.
