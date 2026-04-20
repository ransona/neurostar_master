import ctypes
import re
import time
from dataclasses import dataclass


user32 = ctypes.WinDLL("user32", use_last_error=True)

WM_COMMAND = 0x0111
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT = 0x000C
WM_CLOSE = 0x0010
BM_CLICK = 0x00F5
CB_GETCOUNT = 0x0146
CB_GETCURSEL = 0x0147
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149
CB_SETCURSEL = 0x014E
CB_FINDSTRINGEXACT = 0x0158
CBN_SELCHANGE = 1

TARGET_AP_ID = 1147
TARGET_ML_ID = 1148
TARGET_DV_ID = 1149
CURRENT_AP_ID = 1144
CURRENT_ML_ID = 1145
CURRENT_DV_ID = 1146
GOTO_ID = 1014
STOP_ID = 1018
GOTO_HOME_ID = 1540
GOTO_WORK_ID = 1541
SHOW_INJECTOMATE_COMMAND_ID = 32815
ACTIVE_DRILL_ID = 1043
REFERENCE_SELECTOR_ID = 1387
STEP_AP_ID = 1132
STEP_ML_ID = 1133
STEP_DV_ID = 1134
BUTTON_AP_NEGATIVE_ID = 1103
BUTTON_AP_POSITIVE_ID = 1102
BUTTON_ML_NEGATIVE_ID = 1104
BUTTON_ML_POSITIVE_ID = 1105
BUTTON_DV_NEGATIVE_ID = 1106
BUTTON_DV_POSITIVE_ID = 1107
INJECTION_VOLUME_ID = 10001
INJECTION_GOTO_TEXT_ID = 10004
INJECTION_GOTO_BUTTON_ID = 10005
INJECTION_PLUNGER_POSITION_ID = 10017
SYRINGE_TYPE_ID = 10006
SYRINGE_STEP_UP_ID = 10000
SYRINGE_STEP_DOWN_ID = 10002
SET_REFERENCE_BREGMA_COMMAND_ID = 1095
SET_DRILL_TO_BREGMA_COMMAND_ID = 1071
INJECT_BUTTON_ID = 10018
FILL_BUTTON_ID = 10032
CALIBRATE_INJECTOMATE_ID = 10030
CALIBRATE_SCALE_VALUE_ID = 3242
CLOSE_INJECTOMATE_ID = 10031
NUDGE_STEP_OPTIONS_MM = [0.001, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0]


WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

user32.EnumWindows.argtypes = [WNDENUMPROC, ctypes.c_void_p]
user32.EnumWindows.restype = ctypes.c_bool
user32.EnumChildWindows.argtypes = [ctypes.c_void_p, WNDENUMPROC, ctypes.c_void_p]
user32.EnumChildWindows.restype = ctypes.c_bool
user32.GetWindowTextLengthW.argtypes = [ctypes.c_void_p]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
user32.GetClassNameW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int
user32.GetDlgCtrlID.argtypes = [ctypes.c_void_p]
user32.GetDlgCtrlID.restype = ctypes.c_int
user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint32)]
user32.GetWindowThreadProcessId.restype = ctypes.c_uint32
user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
user32.SendMessageW.restype = ctypes.c_ssize_t


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(RECT)]
user32.GetWindowRect.restype = ctypes.c_bool
@dataclass
class ChildControl:
    hwnd: int
    control_id: int
    class_name: str
    text: str
    left: int
    top: int
    right: int
    bottom: int


class StereoDriveError(RuntimeError):
    pass


class StereoDriveController:
    def __init__(self) -> None:
        self.main_hwnd = self._find_main_window()

    def _refresh_main_window(self) -> None:
        self.main_hwnd = self._find_main_window()

    def get_main_window_rect(self) -> tuple[int, int, int, int]:
        self._refresh_main_window()
        rect = RECT()
        if not user32.GetWindowRect(self.main_hwnd, ctypes.byref(rect)):
            raise StereoDriveError("Could not read StereoDrive window rectangle.")
        return (rect.left, rect.top, max(1, rect.right - rect.left), max(1, rect.bottom - rect.top))

    def get_main_window_handle(self) -> int:
        self._refresh_main_window()
        return int(self.main_hwnd)

    def _find_main_window(self) -> int:
        matches: list[int] = []

        @WNDENUMPROC
        def callback(hwnd: int, _lparam: int) -> bool:
            title = self._window_text(hwnd)
            cls = self._class_name(hwnd)
            if "StereoDrive" in title and cls == "StereoDriveclass":
                matches.append(hwnd)
            return True

        user32.EnumWindows(callback, 0)
        if not matches:
            raise StereoDriveError("StereoDrive main window was not found. Start StereoDrive first.")
        return matches[0]

    def _window_text(self, hwnd: int) -> str:
        length = user32.GetWindowTextLengthW(hwnd)
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, len(buffer))
        return buffer.value

    def _class_name(self, hwnd: int) -> str:
        buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buffer, len(buffer))
        return buffer.value

    def _control_map(self) -> dict[int, int]:
        controls: dict[int, int] = {}

        @WNDENUMPROC
        def callback(hwnd: int, _lparam: int) -> bool:
            ctrl_id = user32.GetDlgCtrlID(hwnd)
            if ctrl_id > 0:
                controls[ctrl_id] = hwnd
            return True

        user32.EnumChildWindows(self.main_hwnd, callback, 0)
        return controls

    def _child_controls(self) -> list[ChildControl]:
        return self._child_controls_for_window(self.main_hwnd)

    def _child_controls_for_window(self, parent_hwnd: int) -> list[ChildControl]:
        controls: list[ChildControl] = []

        @WNDENUMPROC
        def callback(hwnd: int, _lparam: int) -> bool:
            rect = RECT()
            if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return True
            controls.append(
                ChildControl(
                    hwnd=hwnd,
                    control_id=user32.GetDlgCtrlID(hwnd),
                    class_name=self._class_name(hwnd),
                    text=self._get_text(hwnd),
                    left=rect.left,
                    top=rect.top,
                    right=rect.right,
                    bottom=rect.bottom,
                )
            )
            return True

        user32.EnumChildWindows(parent_hwnd, callback, 0)
        return controls

    def _window_process_id(self, hwnd: int) -> int:
        process_id = ctypes.c_uint32(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        return int(process_id.value)

    def _top_level_windows_for_process(self) -> list[ChildControl]:
        self._refresh_main_window()
        wanted_pid = self._window_process_id(self.main_hwnd)
        return [window for window in self._top_level_windows() if self._window_process_id(window.hwnd) == wanted_pid]

    def _top_level_windows(self) -> list[ChildControl]:
        windows: list[ChildControl] = []

        @WNDENUMPROC
        def callback(hwnd: int, _lparam: int) -> bool:
            rect = RECT()
            if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return True
            windows.append(
                ChildControl(
                    hwnd=hwnd,
                    control_id=user32.GetDlgCtrlID(hwnd),
                    class_name=self._class_name(hwnd),
                    text=self._window_text(hwnd).strip(),
                    left=rect.left,
                    top=rect.top,
                    right=rect.right,
                    bottom=rect.bottom,
                )
            )
            return True

        user32.EnumWindows(callback, 0)
        return windows

    def _control_row(self, control: ChildControl) -> dict[str, object]:
        return {
            "handle": control.hwnd,
            "process_id": self._window_process_id(control.hwnd),
            "control_id": control.control_id,
            "class_name": control.class_name,
            "text": control.text,
            "left": control.left,
            "top": control.top,
            "right": control.right,
            "bottom": control.bottom,
        }

    def _numeric_candidates_from_rows(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        numeric_candidates: list[dict[str, object]] = []
        for row in rows:
            text = str(row.get("text", ""))
            for match in re.finditer(r"-?\d+(?:\.\d+)?", text.replace(",", "")):
                value = float(match.group(0))
                if 0.0 <= value <= 5000.0:
                    numeric_candidates.append(row | {"value": value})
        return numeric_candidates

    def _control_handle(self, control_id: int, timeout_seconds: float = 5.0, poll_seconds: float = 0.2) -> int:
        deadline = time.monotonic() + timeout_seconds
        last_seen_controls: list[int] = []
        while time.monotonic() < deadline:
            self._refresh_main_window()
            controls = self._control_map()
            hwnd = controls.get(control_id)
            if hwnd:
                return hwnd
            last_seen_controls = sorted(controls.keys())
            time.sleep(poll_seconds)
        if last_seen_controls:
            preview = ", ".join(str(control) for control in last_seen_controls[:12])
            raise StereoDriveError(
                f"Control ID {control_id} was not found in StereoDrive. Visible controls included: {preview}"
            )
        raise StereoDriveError(f"Control ID {control_id} was not found in StereoDrive.")

    def _child_control_handle(self, parent_hwnd: int, control_id: int) -> int | None:
        match: list[int] = []

        @WNDENUMPROC
        def callback(hwnd: int, _lparam: int) -> bool:
            if user32.GetDlgCtrlID(hwnd) == control_id:
                match.append(hwnd)
                return False
            return True

        user32.EnumChildWindows(parent_hwnd, callback, 0)
        return match[0] if match else None

    def _get_text(self, hwnd: int) -> str:
        length = int(user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0))
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.SendMessageW(hwnd, WM_GETTEXT, len(buffer), ctypes.cast(buffer, ctypes.c_void_p))
        return buffer.value.strip()

    def _set_text(self, hwnd: int, text: str) -> None:
        text_buffer = ctypes.create_unicode_buffer(text)
        user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.cast(text_buffer, ctypes.c_void_p))

    def _notify_command(self, control_id: int, notify_code: int, hwnd: int) -> None:
        wparam = (notify_code << 16) | (control_id & 0xFFFF)
        user32.SendMessageW(self.main_hwnd, WM_COMMAND, wparam, hwnd)

    def _send_command(self, command_id: int) -> None:
        user32.SendMessageW(self.main_hwnd, WM_COMMAND, command_id, 0)

    def _combo_select_exact(self, control_id: int, text: str) -> None:
        hwnd = self._control_handle(control_id)
        text_buffer = ctypes.create_unicode_buffer(text)
        match_index = user32.SendMessageW(hwnd, CB_FINDSTRINGEXACT, -1, ctypes.cast(text_buffer, ctypes.c_void_p))
        if match_index < 0:
            raise StereoDriveError(f"Could not find combo entry '{text}' in control {control_id}.")
        selected_index = user32.SendMessageW(hwnd, CB_SETCURSEL, match_index, 0)
        if selected_index < 0:
            raise StereoDriveError(f"Failed to select combo entry '{text}' in control {control_id}.")
        self._notify_command(control_id, CBN_SELCHANGE, hwnd)

    def _combo_selected_text(self, control_id: int) -> str:
        hwnd = self._control_handle(control_id)
        selected_index = user32.SendMessageW(hwnd, CB_GETCURSEL, 0, 0)
        if selected_index < 0:
            return self._get_text(hwnd)
        text_length = user32.SendMessageW(hwnd, CB_GETLBTEXTLEN, selected_index, 0)
        if text_length < 0:
            return self._get_text(hwnd)
        buffer = ctypes.create_unicode_buffer(text_length + 1)
        user32.SendMessageW(hwnd, CB_GETLBTEXT, selected_index, ctypes.cast(buffer, ctypes.c_void_p))
        return buffer.value.strip()

    def _click(self, control_id: int) -> None:
        hwnd = self._control_handle(control_id)
        user32.SendMessageW(hwnd, BM_CLICK, 0, 0)
        user32.SendMessageW(self.main_hwnd, WM_COMMAND, control_id, hwnd)

    def _parse_float(self, control_id: int) -> float:
        hwnd = self._control_handle(control_id)
        text = self._get_text(hwnd)
        if not text:
            raise StereoDriveError(f"Control ID {control_id} has no numeric value.")
        return float(text)

    def _axis_ids(self, axis: str) -> tuple[int, int, int]:
        normalized = axis.upper()
        mapping = {
            "AP": (CURRENT_AP_ID, STEP_AP_ID, BUTTON_AP_POSITIVE_ID),
            "ML": (CURRENT_ML_ID, STEP_ML_ID, BUTTON_ML_POSITIVE_ID),
            "DV": (CURRENT_DV_ID, STEP_DV_ID, BUTTON_DV_POSITIVE_ID),
        }
        if normalized not in mapping:
            raise StereoDriveError(f"Unknown axis '{axis}'.")
        return mapping[normalized]

    def _axis_button_ids(self, axis: str) -> tuple[int, int]:
        normalized = axis.upper()
        mapping = {
            "AP": (BUTTON_AP_NEGATIVE_ID, BUTTON_AP_POSITIVE_ID),
            "ML": (BUTTON_ML_NEGATIVE_ID, BUTTON_ML_POSITIVE_ID),
            "DV": (BUTTON_DV_NEGATIVE_ID, BUTTON_DV_POSITIVE_ID),
        }
        if normalized not in mapping:
            raise StereoDriveError(f"Unknown axis '{axis}'.")
        return mapping[normalized]

    def get_reference_selector(self) -> str:
        try:
            return self._get_text(self._control_handle(REFERENCE_SELECTOR_ID))
        except StereoDriveError:
            return ""

    def injectomate_visible(self) -> bool:
        return INJECTION_VOLUME_ID in self._control_map()

    def show_injectomate(self) -> None:
        if not self.injectomate_visible():
            self._send_command(SHOW_INJECTOMATE_COMMAND_ID)
            time.sleep(0.3)

    def set_injection_volume(self, volume_label: str) -> None:
        self.show_injectomate()
        self._combo_select_exact(INJECTION_VOLUME_ID, volume_label)
        actual = self._combo_selected_text(INJECTION_VOLUME_ID)
        if actual != volume_label:
            raise StereoDriveError(f"Failed to set injection volume to {volume_label}. Got '{actual}'.")

    def get_injection_volume(self) -> str:
        self.show_injectomate()
        return self._combo_selected_text(INJECTION_VOLUME_ID)

    def set_syringe_type(self, syringe_label: str) -> None:
        self.show_injectomate()
        self._combo_select_exact(SYRINGE_TYPE_ID, syringe_label)
        actual = self._combo_selected_text(SYRINGE_TYPE_ID)
        if actual != syringe_label:
            raise StereoDriveError(f"Failed to set syringe type to {syringe_label}. Got '{actual}'.")

    def inject(self) -> None:
        self.show_injectomate()
        self._click(INJECT_BUTTON_ID)

    def syringe_step_up(self) -> None:
        self.show_injectomate()
        self._click(SYRINGE_STEP_UP_ID)

    def syringe_step_down(self) -> None:
        self.show_injectomate()
        self._click(SYRINGE_STEP_DOWN_ID)

    def syringe_step(self, volume_label: str, up: bool = True) -> None:
        self.set_injection_volume(volume_label)
        if up:
            self.syringe_step_up()
        else:
            self.syringe_step_down()

    def empty_syringe(self) -> None:
        self.show_injectomate()
        hwnd = self._control_handle(INJECTION_GOTO_TEXT_ID)
        self._set_text(hwnd, "0")
        time.sleep(0.1)
        self._click(INJECTION_GOTO_BUTTON_ID)

    def get_injection_plunger_position_nl(self) -> float | None:
        if not self.injectomate_visible():
            return None
        candidates = self.get_injection_numeric_readouts()
        if not candidates:
            return None
        gauge_candidates = [candidate for candidate in candidates if candidate["control_id"] <= 0]
        if gauge_candidates:
            return float(max(gauge_candidates, key=lambda candidate: candidate["top"])["value"])
        return float(max(candidates, key=lambda candidate: candidate["top"])["value"])

    def get_injection_numeric_readouts(self) -> list[dict[str, float | int | str]]:
        all_controls = self._child_controls()
        injection_controls = [control for control in all_controls if control.control_id >= 10000]
        if not injection_controls:
            return []
        left = min(control.left for control in injection_controls)
        right = max(control.right for control in injection_controls)
        top = min(control.top for control in injection_controls)
        bottom = max(control.bottom for control in injection_controls)
        candidates: list[dict[str, float | int | str]] = []
        for control in all_controls:
            if control.right < left - 20 or control.left > right + 20 or control.bottom < top - 20 or control.top > bottom + 20:
                continue
            if control.control_id in {INJECTION_VOLUME_ID, INJECTION_GOTO_TEXT_ID, SYRINGE_TYPE_ID}:
                continue
            if control.control_id in {CURRENT_AP_ID, CURRENT_ML_ID, CURRENT_DV_ID, TARGET_AP_ID, TARGET_ML_ID, TARGET_DV_ID}:
                continue
            text = control.text.strip()
            if not text or re.search(r"[A-Za-z]", text):
                continue
            match = re.fullmatch(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
            if not match:
                continue
            value = float(match.group(0))
            if 0.0 <= value <= 5000.0:
                candidates.append(
                    {
                        "control_id": control.control_id,
                        "class_name": control.class_name,
                        "text": text,
                        "value": value,
                        "left": control.left,
                        "top": control.top,
                        "right": control.right,
                        "bottom": control.bottom,
                    }
                )
        return candidates

    def get_mmc_depth_gauge_rect(self) -> tuple[int, int, int, int] | None:
        gauge = self._mmc_depth_gauge_control()
        if gauge is None:
            return None
        width = max(1, gauge.right - gauge.left)
        height = max(1, gauge.bottom - gauge.top)
        return (gauge.left, gauge.top, width, height)

    def get_mmc_depth_gauge_handle(self) -> int | None:
        gauge = self._mmc_depth_gauge_control()
        if gauge is None:
            return None
        return int(gauge.hwnd)

    def _mmc_depth_gauge_control(self) -> ChildControl | None:
        gauge_controls = [
            control
            for control in self._child_controls()
            if control.class_name == "AfxWnd140" and control.text == "MMCDepth"
        ]
        if not gauge_controls:
            return None
        return max(gauge_controls, key=lambda control: (control.bottom - control.top) * (control.right - control.left))

    def set_current_location_to_bregma(self) -> None:
        self._click(ACTIVE_DRILL_ID)
        time.sleep(0.1)
        self._send_command(SET_REFERENCE_BREGMA_COMMAND_ID)
        time.sleep(0.1)
        self._send_command(SET_DRILL_TO_BREGMA_COMMAND_ID)

    def goto_home(self) -> None:
        self._click(GOTO_HOME_ID)

    def goto_work(self) -> None:
        self._click(GOTO_WORK_ID)

    def fill_injectomate(self) -> None:
        self.show_injectomate()
        self._click(FILL_BUTTON_ID)

    def open_injectomate_calibrate(self) -> None:
        self.show_injectomate()
        self._click(CALIBRATE_INJECTOMATE_ID)
        time.sleep(0.8)

    def _find_injectomate_calibrate_window(self, timeout_seconds: float = 2.0) -> int:
        deadline = time.monotonic() + timeout_seconds
        wanted_pid = self._window_process_id(self.main_hwnd)
        while time.monotonic() < deadline:
            for window in self._top_level_windows():
                if self._window_process_id(window.hwnd) != wanted_pid:
                    continue
                if "Microdrive Calibrate Scale - Injectomate" in window.text:
                    return window.hwnd
                if self._child_control_handle(window.hwnd, CALIBRATE_SCALE_VALUE_ID):
                    return window.hwnd
            time.sleep(0.1)
        raise StereoDriveError("Injectomate calibrate scale popup was not found.")

    def read_injectomate_calibrate_scale_nl(self, close_popup: bool = True) -> float:
        self.open_injectomate_calibrate()
        popup_hwnd = self._find_injectomate_calibrate_window()
        value_hwnd = self._child_control_handle(popup_hwnd, CALIBRATE_SCALE_VALUE_ID)
        if not value_hwnd:
            raise StereoDriveError(f"Calibrate scale value control ID {CALIBRATE_SCALE_VALUE_ID} was not found.")
        text = self._get_text(value_hwnd)
        if close_popup:
            user32.SendMessageW(popup_hwnd, WM_CLOSE, 0, 0)
        try:
            return float(text)
        except ValueError as exc:
            raise StereoDriveError(f"Calibrate scale value was not numeric: '{text}'.") from exc

    def get_injectomate_calibrate_snapshot(self) -> dict[str, object]:
        before_windows = [self._control_row(window) for window in self._top_level_windows()]
        self.open_injectomate_calibrate()
        after_windows = [window for window in self._top_level_windows() if window.right > window.left and window.bottom > window.top]
        windows = [
            window
            for window in after_windows
            if window.hwnd != self.main_hwnd
            and (
                self._window_process_id(window.hwnd) == self._window_process_id(self.main_hwnd)
                or window.class_name in {"#32770", "Afx:00400000:0"}
                or "calib" in window.text.lower()
                or "inject" in window.text.lower()
                or "syringe" in window.text.lower()
            )
        ]
        window_rows: list[dict[str, object]] = []
        all_child_rows: list[dict[str, object]] = []
        for window in windows:
            child_rows: list[dict[str, object]] = []
            for control in self._child_controls_for_window(window.hwnd):
                row = self._control_row(control)
                child_rows.append(row)
                all_child_rows.append(row)
            window_rows.append(
                {
                    "handle": window.hwnd,
                    "process_id": self._window_process_id(window.hwnd),
                    "class_name": window.class_name,
                    "text": window.text,
                    "left": window.left,
                    "top": window.top,
                    "right": window.right,
                    "bottom": window.bottom,
                    "children": child_rows,
                }
            )
        main_child_rows = [self._control_row(control) for control in self._child_controls()]
        all_rows = all_child_rows + main_child_rows
        return {
            "before_windows": before_windows,
            "after_windows": [self._control_row(window) for window in after_windows],
            "candidate_windows": window_rows,
            "main_children_after": main_child_rows,
            "numeric_candidates": self._numeric_candidates_from_rows(all_rows),
        }

    def close_injectomate(self) -> None:
        self._click(CLOSE_INJECTOMATE_ID)

    def get_current_position(self) -> tuple[float, float, float]:
        return (
            self._parse_float(CURRENT_AP_ID),
            self._parse_float(CURRENT_ML_ID),
            self._parse_float(CURRENT_DV_ID),
        )

    def get_current_axis(self, axis: str) -> float:
        current_id, _step_id, _positive_id = self._axis_ids(axis)
        return self._parse_float(current_id)

    def set_nudge_step(self, axis: str, step_mm: float) -> None:
        _current_id, step_id, _positive_id = self._axis_ids(axis)
        label = self._format_step_label(step_mm)
        self._combo_select_exact(step_id, label)
        actual = self._combo_selected_text(step_id)
        if actual != label:
            raise StereoDriveError(f"Failed to set {axis.upper()} nudge step to {label}. Got '{actual}'.")

    def _format_step_label(self, step_mm: float) -> str:
        if step_mm >= 1.0 and float(step_mm).is_integer():
            return f"{int(step_mm)} mm"
        trimmed = f"{step_mm:.3f}".rstrip("0").rstrip(".")
        return f"{trimmed} mm"

    def choose_nudge_step(self, remaining_distance_mm: float, max_step_mm: float | None = None) -> float:
        remaining = abs(remaining_distance_mm)
        if max_step_mm is not None:
            remaining = min(remaining, max_step_mm)
        candidates = [step for step in NUDGE_STEP_OPTIONS_MM if step <= remaining + 1e-9]
        if candidates:
            return candidates[-1]
        return NUDGE_STEP_OPTIONS_MM[0]

    def nudge_axis(self, axis: str, positive: bool) -> None:
        negative_button_id, positive_button_id = self._axis_button_ids(axis)
        self._click(positive_button_id if positive else negative_button_id)

    def move_axis_to_target(
        self,
        axis: str,
        target: float,
        step_mm: float = 5.0,
        tolerance: float = 0.003,
        stop_requested=None,
        status_callback=None,
        dwell_seconds: float = 0.02,
    ) -> None:
        max_iterations = 10000
        moved = False
        positive = False
        active_step: float | None = None
        for _ in range(max_iterations):
            if stop_requested is not None and stop_requested():
                raise StereoDriveError("Operation paused.")
            current = self.get_current_axis(axis)
            diff = target - current
            if abs(diff) <= tolerance:
                return
            chosen_step = self.choose_nudge_step(diff, max_step_mm=step_mm)
            if active_step is None or not abs(active_step - chosen_step) < 1e-9:
                self.set_nudge_step(axis, chosen_step)
                active_step = chosen_step
            if not moved:
                positive = diff > 0
            elif (positive and current >= target) or ((not positive) and current <= target):
                return
            self.nudge_axis(axis, positive)
            moved = True
            if status_callback is not None:
                status_callback(f"Nudging {axis.upper()} to {target:.3f} (current {current:.3f})")
            time.sleep(dwell_seconds)
        raise StereoDriveError(f"Timed out moving {axis.upper()} to target {target:.3f}.")

    def move_to_position_nudged(
        self,
        ap: float,
        ml: float,
        dv: float,
        step_mm: float = 0.005,
        stop_requested=None,
        status_callback=None,
        dwell_seconds: float = 0.02,
    ) -> None:
        self.move_planar_to_target(
            ap,
            ml,
            step_mm=step_mm,
            stop_requested=stop_requested,
            status_callback=status_callback,
            dwell_seconds=dwell_seconds,
        )
        self.move_axis_to_target("DV", dv, step_mm=step_mm, stop_requested=stop_requested, status_callback=status_callback, dwell_seconds=dwell_seconds)

    def move_planar_to_target(
        self,
        ap: float,
        ml: float,
        step_mm: float = 5.0,
        tolerance: float = 0.003,
        stop_requested=None,
        status_callback=None,
        dwell_seconds: float = 0.02,
    ) -> None:
        max_iterations = 20000
        active_steps: dict[str, float] = {}
        move_directions: dict[str, bool] = {}
        moved_axes: dict[str, bool] = {"AP": False, "ML": False}
        for _ in range(max_iterations):
            if stop_requested is not None and stop_requested():
                raise StereoDriveError("Operation paused.")
            current_ap, current_ml, _current_dv = self.get_current_position()
            diffs = {"AP": ap - current_ap, "ML": ml - current_ml}
            remaining_axes = [axis for axis, diff in diffs.items() if abs(diff) > tolerance]
            if not remaining_axes:
                return

            def can_continue(axis: str) -> bool:
                current_value = current_ap if axis == "AP" else current_ml
                target_value = ap if axis == "AP" else ml
                if not moved_axes[axis]:
                    return True
                direction_positive = move_directions[axis]
                if direction_positive and current_value >= target_value:
                    return False
                if (not direction_positive) and current_value <= target_value:
                    return False
                return True

            candidate_axes = [axis for axis in remaining_axes if can_continue(axis)]
            if not candidate_axes:
                return
            axis = max(candidate_axes, key=lambda name: abs(diffs[name]))
            diff = diffs[axis]
            chosen_step = self.choose_nudge_step(diff, max_step_mm=step_mm)
            previous_step = active_steps.get(axis)
            if previous_step is None or not abs(previous_step - chosen_step) < 1e-9:
                self.set_nudge_step(axis, chosen_step)
                active_steps[axis] = chosen_step
            if not moved_axes[axis]:
                move_directions[axis] = diff > 0
            self.nudge_axis(axis, move_directions[axis])
            moved_axes[axis] = True
            if status_callback is not None:
                status_callback(f"Nudging XY toward [{ap:.3f}, {ml:.3f}]")
            time.sleep(dwell_seconds)
        raise StereoDriveError(f"Timed out moving AP/ML to target [{ap:.3f}, {ml:.3f}].")

    def set_target_position(self, ap: float, ml: float, dv: float) -> None:
        self._set_text(self._control_handle(TARGET_AP_ID), f"{ap:.2f}")
        self._set_text(self._control_handle(TARGET_ML_ID), f"{ml:.2f}")
        self._set_text(self._control_handle(TARGET_DV_ID), f"{dv:.2f}")
        self._verify_target_position(ap, ml, dv)

    def _verify_target_position(self, ap: float, ml: float, dv: float) -> None:
        actual_ap = self._parse_float(TARGET_AP_ID)
        actual_ml = self._parse_float(TARGET_ML_ID)
        actual_dv = self._parse_float(TARGET_DV_ID)
        if round(actual_ap, 2) != round(ap, 2):
            raise StereoDriveError(f"Failed to set Bregma AP target box to {ap:.2f}.")
        if round(actual_ml, 2) != round(ml, 2):
            raise StereoDriveError(f"Failed to set Bregma ML target box to {ml:.2f}.")
        if round(actual_dv, 2) != round(dv, 2):
            raise StereoDriveError(f"Failed to set Bregma DV target box to {dv:.2f}.")

    def goto_position(self, ap: float, ml: float, dv: float, delay_seconds: float = 0.5) -> None:
        self.set_target_position(ap, ml, dv)
        time.sleep(delay_seconds)
        self._click(GOTO_ID)

    def wait_for_position(
        self,
        ap: float,
        ml: float,
        dv: float,
        tolerance_mm: float = 0.02,
        timeout_seconds: float = 60.0,
        poll_seconds: float = 0.1,
        stop_requested=None,
    ) -> None:
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            if stop_requested is not None and stop_requested():
                raise StereoDriveError("Operation paused.")
            current_ap, current_ml, current_dv = self.get_current_position()
            if (
                abs(current_ap - ap) <= tolerance_mm
                and abs(current_ml - ml) <= tolerance_mm
                and abs(current_dv - dv) <= tolerance_mm
            ):
                return
            time.sleep(poll_seconds)
        raise StereoDriveError(
            f"Timed out waiting for position [{ap:.2f}, {ml:.2f}, {dv:.2f}] in StereoDrive."
        )

    def stop(self) -> None:
        self._click(STOP_ID)
