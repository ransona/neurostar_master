import ctypes
import time


user32 = ctypes.WinDLL("user32", use_last_error=True)

WM_COMMAND = 0x0111
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT = 0x000C
BM_CLICK = 0x00F5

TARGET_AP_ID = 1147
TARGET_ML_ID = 1148
TARGET_DV_ID = 1149
CURRENT_AP_ID = 1144
CURRENT_ML_ID = 1145
CURRENT_DV_ID = 1146
GOTO_ID = 1014
STOP_ID = 1018
REFERENCE_SELECTOR_ID = 1387


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
