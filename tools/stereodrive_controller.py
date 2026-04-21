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
MN_GETHMENU = 0x01E1
MF_BYPOSITION = 0x00000400
BM_CLICK = 0x00F5
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
CB_GETCOUNT = 0x0146
CB_GETCURSEL = 0x0147
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149
CB_SETCURSEL = 0x014E
CB_FINDSTRINGEXACT = 0x0158
CBN_SELCHANGE = 1
EN_CHANGE = 0x0300
EN_UPDATE = 0x0400

TARGET_AP_ID = 1147
TARGET_ML_ID = 1148
TARGET_DV_ID = 1149
CURRENT_AP_ID = 1144
CURRENT_ML_ID = 1145
CURRENT_DV_ID = 1146
GOTO_ID = 1014
STOP_ID = 1018
TOOLS_BUTTON_ID = 1010
GOTO_HOME_ID = 1540
GOTO_WORK_ID = 1541
SHOW_INJECTOMATE_COMMAND_ID = 32815
SHOW_REFERENCE_PANEL_COMMAND_ID = 32809
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
INJECTION_STATUS_RATE_ID = 10028
INJECTION_STATUS_TIME_ELAPSED_ID = 10020
INJECTION_STATUS_TIME_REMAINING_ID = 10021
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
CLOSE_REFERENCE_PANEL_ID = 1042
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
user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
user32.PostMessageW.restype = ctypes.c_bool
user32.IsWindowVisible.argtypes = [ctypes.c_void_p]
user32.IsWindowVisible.restype = ctypes.c_bool
user32.IsWindowEnabled.argtypes = [ctypes.c_void_p]
user32.IsWindowEnabled.restype = ctypes.c_bool
user32.GetMenuItemCount.argtypes = [ctypes.c_void_p]
user32.GetMenuItemCount.restype = ctypes.c_int
user32.GetMenuStringW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_int, ctypes.c_uint]
user32.GetMenuStringW.restype = ctypes.c_int
user32.GetMenuItemID.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.GetMenuItemID.restype = ctypes.c_uint
user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
user32.SetForegroundWindow.restype = ctypes.c_bool
user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
user32.SetCursorPos.restype = ctypes.c_bool
user32.mouse_event.argtypes = [ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p]
user32.mouse_event.restype = None


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
        self._active_injectomate_trigger_control_id: int | None = None

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
        visible_matches = [hwnd for hwnd in matches if user32.IsWindowVisible(hwnd)]
        if visible_matches:
            return visible_matches[0]
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
        self._refresh_main_window()
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

    def _click_popup_ok_button(self, popup_hwnd: int) -> bool:
        for control in self._child_controls_for_window(popup_hwnd):
            if control.class_name != "Button":
                continue
            if re.fullmatch(r"\s*OK\s*", control.text, re.IGNORECASE):
                user32.PostMessageW(control.hwnd, BM_CLICK, 0, 0)
                return True
        return False

    def _click_popup_button(self, popup_hwnd: int, label: str) -> bool:
        wanted = label.strip().lower()
        for control in self._child_controls_for_window(popup_hwnd):
            if control.class_name != "Button":
                continue
            text = control.text.replace("&", "").strip().lower()
            if text == wanted:
                user32.PostMessageW(control.hwnd, BM_CLICK, 0, 0)
                return True
        return False

    def _combined_window_text(self, hwnd: int) -> str:
        parts = [self._window_text(hwnd)]
        parts.extend(control.text for control in self._child_controls_for_window(hwnd) if control.text)
        return "\n".join(parts)

    def _find_below_skull_warning_dialog(self) -> int | None:
        for window in self._top_level_windows_for_process():
            if window.class_name != "#32770":
                continue
            text = self._combined_window_text(window.hwnd).lower()
            if "target position is below the skull surface" in text and "do you want to continue" in text:
                return window.hwnd
        return None

    def _find_no_actual_movement_dialog(self) -> int | None:
        for window in self._top_level_windows_for_process():
            if window.class_name != "#32770":
                continue
            text = self._combined_window_text(window.hwnd).lower()
            if "there are no actual movements to execute" in text:
                return window.hwnd
        return None

    def confirm_below_skull_warning(self, timeout_seconds: float = 0.75, poll_seconds: float = 0.02) -> bool:
        deadline = time.monotonic() + timeout_seconds
        while True:
            popup_hwnd = self._find_below_skull_warning_dialog()
            if popup_hwnd is not None:
                return self._click_popup_button(popup_hwnd, "Yes")
            if time.monotonic() >= deadline:
                break
            time.sleep(poll_seconds)
        return False

    def confirm_no_actual_movement_dialog(self, timeout_seconds: float = 0.5, poll_seconds: float = 0.02) -> bool:
        deadline = time.monotonic() + timeout_seconds
        while True:
            popup_hwnd = self._find_no_actual_movement_dialog()
            if popup_hwnd is not None:
                return self._click_popup_ok_button(popup_hwnd)
            if time.monotonic() >= deadline:
                break
            time.sleep(poll_seconds)
        return False

    def _get_text(self, hwnd: int) -> str:
        length = int(user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0))
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.SendMessageW(hwnd, WM_GETTEXT, len(buffer), ctypes.cast(buffer, ctypes.c_void_p))
        return buffer.value.strip()

    def _window_rect_tuple(self, hwnd: int) -> tuple[int, int, int, int] | None:
        rect = RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return None
        return (rect.left, rect.top, rect.right, rect.bottom)

    def _set_text(self, hwnd: int, text: str) -> None:
        text_buffer = ctypes.create_unicode_buffer(text)
        user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.cast(text_buffer, ctypes.c_void_p))

    def _set_edit_control_text(self, control_id: int, text: str) -> None:
        hwnd = self._control_handle(control_id)
        for _attempt in range(3):
            self._set_text(hwnd, text)
            self._notify_command(control_id, EN_UPDATE, hwnd)
            self._notify_command(control_id, EN_CHANGE, hwnd)
            time.sleep(0.05)
            actual = self._get_text(hwnd)
            if actual == text:
                return
        actual = self._get_text(hwnd)
        raise StereoDriveError(f"Target control {control_id} was not set to '{text}'. Current value: '{actual}'.")

    def _notify_command(self, control_id: int, notify_code: int, hwnd: int) -> None:
        wparam = (notify_code << 16) | (control_id & 0xFFFF)
        user32.SendMessageW(self.main_hwnd, WM_COMMAND, wparam, hwnd)

    def _send_command(self, command_id: int) -> None:
        self._refresh_main_window()
        user32.SendMessageW(
            ctypes.c_void_p(self.main_hwnd),
            WM_COMMAND,
            ctypes.c_void_p(command_id),
            ctypes.c_void_p(0),
        )

    def _normalize_menu_label(self, text: str) -> str:
        normalized = text.lower().replace("&", "").replace("...", "").replace("\u2026", "")
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized.strip()

    def _popup_menu_window(self, timeout_seconds: float = 1.0) -> int:
        self._refresh_main_window()
        wanted_pid = self._window_process_id(self.main_hwnd)
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            for window in self._top_level_windows():
                if window.class_name == "#32768" and self._window_process_id(window.hwnd) == wanted_pid:
                    return window.hwnd
            time.sleep(0.03)
        raise StereoDriveError("StereoDrive Tools popup menu was not found.")

    def _popup_menu_handle(self, popup_hwnd: int) -> int:
        menu = user32.SendMessageW(popup_hwnd, MN_GETHMENU, 0, 0)
        if not menu:
            raise StereoDriveError("Could not retrieve the StereoDrive Tools popup menu handle.")
        return int(menu)

    def _popup_menu_items(self, menu_handle: int) -> list[tuple[int, str, str]]:
        count = user32.GetMenuItemCount(menu_handle)
        if count < 0:
            return []
        items: list[tuple[int, str, str]] = []
        for position in range(count):
            buffer = ctypes.create_unicode_buffer(256)
            user32.GetMenuStringW(menu_handle, position, buffer, len(buffer), MF_BYPOSITION)
            item_id = user32.GetMenuItemID(menu_handle, position)
            if item_id == 0xFFFFFFFF:
                continue
            label = buffer.value
            items.append((int(item_id), label, self._normalize_menu_label(label)))
        return items

    def _find_popup_menu_item_id(self, menu_handle: int, wanted_label: str) -> int | None:
        wanted = self._normalize_menu_label(wanted_label)
        for item_id, _label, normalized in self._popup_menu_items(menu_handle):
            if normalized == wanted:
                return item_id
        return None

    def _find_popup_menu_item_position(self, menu_handle: int, wanted_label: str) -> int | None:
        wanted = self._normalize_menu_label(wanted_label)
        count = user32.GetMenuItemCount(menu_handle)
        if count < 0:
            return None
        for position in range(count):
            buffer = ctypes.create_unicode_buffer(256)
            user32.GetMenuStringW(menu_handle, position, buffer, len(buffer), MF_BYPOSITION)
            if self._normalize_menu_label(buffer.value) == wanted:
                return position
        return None

    def _invoke_tools_menu_item(self, wanted_label: str) -> None:
        self._click(TOOLS_BUTTON_ID)
        try:
            popup_hwnd = self._popup_menu_window(timeout_seconds=0.35)
        except StereoDriveError:
            tools_hwnd = self._control_handle(TOOLS_BUTTON_ID, timeout_seconds=0.5, poll_seconds=0.05)
            rect = self._window_rect_tuple(tools_hwnd)
            if rect is None:
                raise
            left, top, right, bottom = rect
            user32.SetForegroundWindow(self.main_hwnd)
            time.sleep(0.15)
            user32.SetCursorPos(int((left + right) / 2), int((top + bottom) / 2))
            time.sleep(0.08)
            user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, None)
            time.sleep(0.03)
            user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, None)
            popup_hwnd = self._popup_menu_window(timeout_seconds=1.0)
        menu_handle = self._popup_menu_handle(popup_hwnd)
        item_id = self._find_popup_menu_item_id(menu_handle, wanted_label)
        if item_id is None:
            available = ", ".join(label for _item_id, label, _normalized in self._popup_menu_items(menu_handle) if label)
            raise StereoDriveError(
                f"StereoDrive Tools menu item '{wanted_label}' was not found. Available items: {available}"
            )
        self._send_command(item_id)
        time.sleep(0.3)

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

    def click_tools_button(self) -> None:
        self._click(TOOLS_BUTTON_ID)

    def click_synchronize_drill_and_syringe_menu_item(self) -> None:
        self.send_sync_direct_command()

    def send_sync_direct_command(self) -> None:
        self._send_command(SHOW_REFERENCE_PANEL_COMMAND_ID)

    def _click_tools_then_menu_position(self, wanted_label: str) -> None:
        self.click_tools_button()
        popup_hwnd = self._popup_menu_window(timeout_seconds=1.0)
        menu_handle = self._popup_menu_handle(popup_hwnd)
        position = self._find_popup_menu_item_position(menu_handle, wanted_label)
        if position is None:
            available = ", ".join(label for _item_id, label, _normalized in self._popup_menu_items(menu_handle) if label)
            raise StereoDriveError(
                f"StereoDrive Tools menu item '{wanted_label}' was not found. Available items: {available}"
            )
        rect = self._window_rect_tuple(popup_hwnd)
        if rect is None:
            raise StereoDriveError("Could not read StereoDrive Tools popup rectangle.")
        left, top, right, _bottom = rect
        item_height = 24
        x = left + max(20, min(180, int((right - left) / 2)))
        y = top + int((position + 0.5) * item_height)
        user32.SetCursorPos(x, y)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, None)
        time.sleep(0.03)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, None)
        time.sleep(0.3)

    def _is_control_enabled(self, control_id: int) -> bool:
        hwnd = self._control_handle(control_id, timeout_seconds=0.2, poll_seconds=0.01)
        return bool(user32.IsWindowEnabled(hwnd))

    def _post_click(self, control_id: int) -> None:
        hwnd = self._control_handle(control_id)
        user32.PostMessageW(hwnd, BM_CLICK, 0, 0)

    def _post_click_handle(self, hwnd: int, control_id: int) -> None:
        user32.PostMessageW(hwnd, BM_CLICK, 0, 0)
        user32.PostMessageW(self.main_hwnd, WM_COMMAND, control_id, hwnd)

    def _post_command(self, control_id: int) -> None:
        hwnd = self._control_handle(control_id)
        user32.PostMessageW(self.main_hwnd, WM_COMMAND, control_id, hwnd)

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

    def reference_panel_visible(self) -> bool:
        controls = self._control_map()
        return SET_REFERENCE_BREGMA_COMMAND_ID in controls or REFERENCE_SELECTOR_ID in controls

    def _wait_for_reference_panel(self, timeout_seconds: float = 1.0) -> bool:
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            if self.reference_panel_visible():
                return True
            time.sleep(0.05)
        return self.reference_panel_visible()

    def show_reference_panel(self) -> None:
        if self.reference_panel_visible():
            return
        self._send_command(SHOW_REFERENCE_PANEL_COMMAND_ID)

        if self._wait_for_reference_panel(timeout_seconds=3.0):
            return

        raise StereoDriveError("Synchronize Drill and Syringe panel did not open.")

    def close_reference_panel(self) -> None:
        if not self.reference_panel_visible():
            return
        self._click(CLOSE_REFERENCE_PANEL_ID)
        deadline = time.monotonic() + 1.0
        while time.monotonic() < deadline:
            if not self.reference_panel_visible():
                return
            time.sleep(0.05)

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

    def stop_injectomate_motion(self, trigger_control_id: int | None = None) -> None:
        self.show_injectomate()
        controls = self._control_map()
        candidates: list[int] = []
        active_trigger = trigger_control_id or self._active_injectomate_trigger_control_id
        if active_trigger is not None:
            candidates.append(active_trigger)
        candidates.extend([INJECT_BUTTON_ID, SYRINGE_STEP_UP_ID, SYRINGE_STEP_DOWN_ID, INJECTION_GOTO_BUTTON_ID])

        clicked_any = False
        for control_id in dict.fromkeys(candidates):
            hwnd = controls.get(control_id)
            if not hwnd:
                continue
            text = self._get_text(hwnd).strip().lower()
            if control_id == active_trigger or text in {"stop", "cancel"}:
                self._post_click_handle(hwnd, control_id)
                clicked_any = True
        if not clicked_any:
            raise StereoDriveError("No active Injectomate stop control was found.")

    def syringe_step_up(self, stop_requested=None) -> None:
        self.show_injectomate()
        self._click(SYRINGE_STEP_UP_ID)
        self.wait_for_injectomate_motion_complete(SYRINGE_STEP_UP_ID, stop_requested=stop_requested)

    def syringe_step_down(self, stop_requested=None) -> None:
        self.show_injectomate()
        self._click(SYRINGE_STEP_DOWN_ID)
        self.wait_for_injectomate_motion_complete(SYRINGE_STEP_DOWN_ID, stop_requested=stop_requested)

    def syringe_step(self, volume_label: str, up: bool = True, stop_requested=None) -> None:
        self.set_injection_volume(volume_label)
        if up:
            self.syringe_step_up(stop_requested=stop_requested)
        else:
            self.syringe_step_down(stop_requested=stop_requested)

    def empty_syringe(self) -> None:
        self.show_injectomate()
        hwnd = self._control_handle(INJECTION_GOTO_TEXT_ID)
        self._set_text(hwnd, "0")
        time.sleep(0.1)
        self._click(INJECTION_GOTO_BUTTON_ID)
        self.wait_for_injectomate_motion_complete(INJECTION_GOTO_BUTTON_ID, timeout_seconds=180.0)

    def _injectomate_motion_status_text(self) -> str:
        parts = []
        controls = self._control_map()
        for control_id in (INJECTION_STATUS_RATE_ID, INJECTION_STATUS_TIME_ELAPSED_ID, INJECTION_STATUS_TIME_REMAINING_ID):
            hwnd = controls.get(control_id)
            if hwnd:
                text = self._get_text(hwnd).strip()
                if text:
                    parts.append(text)
        return " ".join(parts)

    def wait_for_injectomate_motion_complete(
        self,
        trigger_control_id: int | None = None,
        timeout_seconds: float = 90.0,
        poll_seconds: float = 0.02,
        stop_requested=None,
    ) -> None:
        deadline = time.monotonic() + timeout_seconds
        saw_busy = False
        stable_since: float | None = None
        self._active_injectomate_trigger_control_id = trigger_control_id
        try:
            while time.monotonic() < deadline:
                if stop_requested is not None and stop_requested():
                    try:
                        self.stop_injectomate_motion(trigger_control_id=trigger_control_id)
                    except Exception:
                        pass
                    raise StereoDriveError("Injectomate syringe motion wait stopped.")
                status_busy = bool(self._injectomate_motion_status_text())
                disabled_busy = False
                if trigger_control_id is not None:
                    try:
                        disabled_busy = not self._is_control_enabled(trigger_control_id)
                    except StereoDriveError:
                        disabled_busy = False
                busy = status_busy or disabled_busy
                now = time.monotonic()
                if busy:
                    saw_busy = True
                    stable_since = None
                else:
                    if stable_since is None:
                        stable_since = now
                    required_stable_s = 0.10 if saw_busy else 0.20
                    if now - stable_since >= required_stable_s:
                        return
                time.sleep(poll_seconds)
        finally:
            if self._active_injectomate_trigger_control_id == trigger_control_id:
                self._active_injectomate_trigger_control_id = None
        raise StereoDriveError("Timed out waiting for Injectomate syringe motion to complete.")

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
        self.show_reference_panel()
        try:
            self._click(ACTIVE_DRILL_ID)
            time.sleep(0.1)
            self._send_command(SET_REFERENCE_BREGMA_COMMAND_ID)
            time.sleep(0.1)
            self._send_command(SET_DRILL_TO_BREGMA_COMMAND_ID)
            self._verify_bregma_zeroed()
        finally:
            self.close_reference_panel()

    def _verify_bregma_zeroed(self, timeout_seconds: float = 2.0, tolerance_mm: float = 0.02) -> None:
        deadline = time.monotonic() + timeout_seconds
        last_position: tuple[float, float, float] | None = None
        while time.monotonic() < deadline:
            last_position = self.get_current_position()
            if all(abs(value) <= tolerance_mm for value in last_position):
                return
            time.sleep(0.05)
        if last_position is None:
            raise StereoDriveError("Could not verify Bregma zero after Set Bregma.")
        ap, ml, dv = last_position
        raise StereoDriveError(
            "Set Bregma did not zero the displayed Bregma coordinates. "
            f"Current readings are AP {ap:.3f}, ML {ml:.3f}, DV {dv:.3f} mm."
        )

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
        time.sleep(0.2)

    def _find_injectomate_calibrate_window(self, timeout_seconds: float = 5.0) -> int:
        deadline = time.monotonic() + timeout_seconds
        wanted_pid = self._window_process_id(self.main_hwnd)
        seen_windows: list[str] = []
        while time.monotonic() < deadline:
            for window in self._top_level_windows():
                if self._window_process_id(window.hwnd) != wanted_pid:
                    continue
                if window.right <= window.left or window.bottom <= window.top:
                    continue
                if window.text:
                    seen_windows.append(f"{window.class_name}:{window.text}")
                if "Microdrive Calibrate Scale - Injectomate" in window.text:
                    return window.hwnd
                if self._child_control_handle(window.hwnd, CALIBRATE_SCALE_VALUE_ID):
                    return window.hwnd
            time.sleep(0.001)
        preview = ", ".join(dict.fromkeys(seen_windows[:8]))
        suffix = f" Seen windows: {preview}" if preview else ""
        raise StereoDriveError(f"Injectomate calibrate scale popup was not found.{suffix}")

    def read_injectomate_calibrate_scale_nl(self, close_popup: bool = True) -> float:
        self.show_injectomate()
        self._post_click(CALIBRATE_INJECTOMATE_ID)
        try:
            popup_hwnd = self._find_injectomate_calibrate_window(timeout_seconds=0.5)
        except StereoDriveError:
            self._post_command(CALIBRATE_INJECTOMATE_ID)
            popup_hwnd = self._find_injectomate_calibrate_window()
        deadline = time.monotonic() + 5.0
        last_text = ""
        value_hwnd = 0
        try:
            while time.monotonic() < deadline:
                value_hwnd = self._child_control_handle(popup_hwnd, CALIBRATE_SCALE_VALUE_ID) or 0
                if not value_hwnd:
                    time.sleep(0.001)
                    continue
                last_text = self._get_text(value_hwnd)
                if re.fullmatch(r"-?\d+(?:\.\d+)?", last_text.replace(",", "")):
                    return float(last_text.replace(",", ""))
                time.sleep(0.001)
            popup_rect = self._window_rect_tuple(popup_hwnd)
            if not value_hwnd:
                raise StereoDriveError(
                    f"Calibrate popup was found but value control ID {CALIBRATE_SCALE_VALUE_ID} was not found. "
                    f"Popup={popup_hwnd} rect={popup_rect}."
                )
            raise StereoDriveError(
                f"Calibrate value control was found but did not contain a numeric value. "
                f"Control={value_hwnd} text='{last_text}'."
            )
        finally:
            if close_popup:
                self._click_popup_ok_button(popup_hwnd)

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
            if axis.upper() == "DV" and positive:
                self.confirm_below_skull_warning(timeout_seconds=0.05, poll_seconds=0.01)
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
        self._set_edit_control_text(TARGET_AP_ID, f"{ap:.2f}")
        time.sleep(0.05)
        self._set_edit_control_text(TARGET_ML_ID, f"{ml:.2f}")
        time.sleep(0.05)
        self._set_edit_control_text(TARGET_DV_ID, f"{dv:.2f}")
        time.sleep(0.2)
        self._verify_target_position(ap, ml, dv)

    def _verify_target_position(self, ap: float, ml: float, dv: float) -> None:
        actual_ap_text = self._get_text(self._control_handle(TARGET_AP_ID))
        actual_ml_text = self._get_text(self._control_handle(TARGET_ML_ID))
        actual_dv_text = self._get_text(self._control_handle(TARGET_DV_ID))
        if not actual_ap_text or not actual_ml_text or not actual_dv_text:
            raise StereoDriveError(
                f"Target position boxes must all be populated before GoTo. "
                f"Current values: AP='{actual_ap_text}', ML='{actual_ml_text}', DV='{actual_dv_text}'."
            )
        actual_ap = float(actual_ap_text)
        actual_ml = float(actual_ml_text)
        actual_dv = float(actual_dv_text)
        if round(actual_ap, 2) != round(ap, 2):
            raise StereoDriveError(f"Failed to set Bregma AP target box to {ap:.2f}.")
        if round(actual_ml, 2) != round(ml, 2):
            raise StereoDriveError(f"Failed to set Bregma ML target box to {ml:.2f}.")
        if round(actual_dv, 2) != round(dv, 2):
            raise StereoDriveError(f"Failed to set Bregma DV target box to {dv:.2f}.")

    def goto_position(self, ap: float, ml: float, dv: float, delay_seconds: float = 0.75) -> None:
        self.set_target_position(ap, ml, dv)
        time.sleep(delay_seconds)
        self._click(GOTO_ID)
        self.confirm_below_skull_warning(timeout_seconds=1.0, poll_seconds=0.02)
        self.confirm_no_actual_movement_dialog(timeout_seconds=0.5, poll_seconds=0.02)

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
            self.confirm_below_skull_warning(timeout_seconds=0.01, poll_seconds=0.005)
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
