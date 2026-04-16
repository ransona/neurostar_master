import ctypes
import time


user32 = ctypes.WinDLL("user32", use_last_error=True)

WM_COMMAND = 0x0111
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT = 0x000C
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
user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
user32.SendMessageW.restype = ctypes.c_void_p


class StereoDriveError(RuntimeError):
    pass


class StereoDriveController:
    def __init__(self) -> None:
        self.main_hwnd = self._find_main_window()

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

    def _control_handle(self, control_id: int) -> int:
        controls = self._control_map()
        hwnd = controls.get(control_id)
        if not hwnd:
            raise StereoDriveError(f"Control ID {control_id} was not found in StereoDrive.")
        return hwnd

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

    def _combo_select_exact(self, control_id: int, text: str) -> None:
        hwnd = self._control_handle(control_id)
        text_buffer = ctypes.create_unicode_buffer(text)
        match_index = int(user32.SendMessageW(hwnd, CB_FINDSTRINGEXACT, -1, ctypes.cast(text_buffer, ctypes.c_void_p)))
        if match_index < 0:
            raise StereoDriveError(f"Could not find combo entry '{text}' in control {control_id}.")
        selected_index = int(user32.SendMessageW(hwnd, CB_SETCURSEL, match_index, 0))
        if selected_index < 0:
            raise StereoDriveError(f"Failed to select combo entry '{text}' in control {control_id}.")
        self._notify_command(control_id, CBN_SELCHANGE, hwnd)

    def _combo_selected_text(self, control_id: int) -> str:
        hwnd = self._control_handle(control_id)
        selected_index = int(user32.SendMessageW(hwnd, CB_GETCURSEL, 0, 0))
        if selected_index < 0:
            return self._get_text(hwnd)
        text_length = int(user32.SendMessageW(hwnd, CB_GETLBTEXTLEN, selected_index, 0))
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
        label = f"{step_mm:.3f} mm"
        self._combo_select_exact(step_id, label)
        actual = self._combo_selected_text(step_id)
        if actual != label:
            raise StereoDriveError(f"Failed to set {axis.upper()} nudge step to {label}. Got '{actual}'.")

    def nudge_axis(self, axis: str, positive: bool) -> None:
        negative_button_id, positive_button_id = self._axis_button_ids(axis)
        self._click(positive_button_id if positive else negative_button_id)

    def move_axis_to_target(
        self,
        axis: str,
        target: float,
        step_mm: float = 0.005,
        tolerance: float = 0.003,
        stop_requested=None,
        status_callback=None,
    ) -> None:
        self.set_nudge_step(axis, step_mm)
        max_iterations = 10000
        for _ in range(max_iterations):
            if stop_requested is not None and stop_requested():
                raise StereoDriveError("Operation paused.")
            current = self.get_current_axis(axis)
            diff = target - current
            if abs(diff) <= tolerance:
                return
            positive = diff > 0
            self.nudge_axis(axis, positive)
            if status_callback is not None:
                status_callback(f"Nudging {axis.upper()} to {target:.3f} (current {current:.3f})")
            time.sleep(0.02)
        raise StereoDriveError(f"Timed out moving {axis.upper()} to target {target:.3f}.")

    def move_to_position_nudged(
        self,
        ap: float,
        ml: float,
        dv: float,
        step_mm: float = 0.005,
        stop_requested=None,
        status_callback=None,
    ) -> None:
        self.move_axis_to_target("AP", ap, step_mm=step_mm, stop_requested=stop_requested, status_callback=status_callback)
        self.move_axis_to_target("ML", ml, step_mm=step_mm, stop_requested=stop_requested, status_callback=status_callback)
        self.move_axis_to_target("DV", dv, step_mm=step_mm, stop_requested=stop_requested, status_callback=status_callback)

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

    def stop(self) -> None:
        self._click(STOP_ID)
