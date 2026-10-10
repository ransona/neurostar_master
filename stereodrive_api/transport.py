"""Exclusive Windows USB virtual serial connection and hardware-free simulator."""
import ctypes
from ctypes import wintypes
import os
import json
from pathlib import Path
import struct
import subprocess
import time
from .protocol import AXES, SELECTORS, Status

DEVICE_SERIAL = '206334AC5031'
DEVICE_ID = 'USB\\VID_0483&PID_5743\\' + DEVICE_SERIAL

class DCB(ctypes.Structure):
    _fields_ = [('length', wintypes.DWORD), ('baud', wintypes.DWORD),
        ('flags', wintypes.DWORD), ('reserved', wintypes.WORD),
        ('xon_limit', wintypes.WORD), ('xoff_limit', wintypes.WORD),
        ('byte_size', wintypes.BYTE), ('parity', wintypes.BYTE),
        ('stop_bits', wintypes.BYTE), ('xon', ctypes.c_char),
        ('xoff', ctypes.c_char), ('error', ctypes.c_char),
        ('eof', ctypes.c_char), ('event', ctypes.c_char), ('reserved1', wintypes.WORD)]

class Timeouts(ctypes.Structure):
    _fields_ = [(name, wintypes.DWORD) for name in
        ('read_interval', 'read_multiplier', 'read_constant', 'write_multiplier', 'write_constant')]

def find_device():
    """Follow the exact USB serial identity, never assume a COM number."""
    if os.name != 'nt':
        raise RuntimeError('Live USB connection requires Windows.')
    import winreg
    key = r'SYSTEM\CurrentControlSet\Enum\USB\VID_0483&PID_5743' + '\\' + DEVICE_SERIAL + r'\Device Parameters'
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as handle:
            port = winreg.QueryValueEx(handle, 'PortName')[0]
    except FileNotFoundError as exc:
        raise RuntimeError('The identified StereoDrive USB controller was not found.') from exc
    if not isinstance(port, str) or not port.startswith('COM') or not port[3:].isdigit():
        raise RuntimeError('Controller has no valid USB serial port assignment.')
    return port

class WindowsSerial:
    def __init__(self, log):
        self.log = log
        self.handle = None
        self.port = find_device()
        native = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq StereoDrive.exe',
                                 '/FO', 'CSV', '/NH'], capture_output=True, text=True, check=True).stdout
        if 'stereodrive.exe' in native.lower():
            raise RuntimeError('Close StereoDrive and its automation/probe before connecting directly.')
        self.identity = DEVICE_ID
        self.k = ctypes.WinDLL('kernel32', use_last_error=True)
        self.k.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
            ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        self.k.CreateFileW.restype = wintypes.HANDLE
        self.k.ReadFile.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                                   ctypes.POINTER(wintypes.DWORD), ctypes.c_void_p]
        self.k.WriteFile.argtypes = self.k.ReadFile.argtypes
        self.k.GetCommState.argtypes = [wintypes.HANDLE, ctypes.POINTER(DCB)]
        self.k.GetCommTimeouts.argtypes = [wintypes.HANDLE, ctypes.POINTER(Timeouts)]
        self.k.SetCommTimeouts.argtypes = self.k.GetCommTimeouts.argtypes
        self.k.CloseHandle.argtypes = [wintypes.HANDLE]
        self.handle = self.k.CreateFileW('\\\\.\\' + self.port, 0xc0000000, 0, None, 3, 0, None)
        if self.handle == ctypes.c_void_p(-1).value:
            self.handle = None
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            dcb = DCB(); dcb.length = ctypes.sizeof(DCB)
            if not self.k.GetCommState(self.handle, ctypes.byref(dcb)):
                raise ctypes.WinError(ctypes.get_last_error())
            self.description = f'{self.port}  /  {dcb.baud} baud  /  {DEVICE_SERIAL}'
            # Retain the driver's DCB. Do not guess baud or toggle DTR/RTS.
            self.original_timeouts = Timeouts()
            if not self.k.GetCommTimeouts(self.handle, ctypes.byref(self.original_timeouts)):
                raise ctypes.WinError(ctypes.get_last_error())
            limits = Timeouts(20, 0, 100, 0, 300)
            if not self.k.SetCommTimeouts(self.handle, ctypes.byref(limits)):
                raise ctypes.WinError(ctypes.get_last_error())
        except Exception:
            self.k.CloseHandle(self.handle); self.handle = None
            raise

    def write(self, data):
        try: self.log('OUT', hex=data.hex())
        except Exception:
            if data[:2] != b'\xaf\x0f' and data != b'\xaf\x11\x00': raise
        sent = wintypes.DWORD(); buf = ctypes.create_string_buffer(data)
        if not self.k.WriteFile(self.handle, buf, len(data), ctypes.byref(sent), None) or sent.value != len(data):
            raise RuntimeError('Serial write failed or was partial; no retransmission.')

    def exchange(self, data, size):
        self.write(data)
        reply = bytearray(); deadline = time.monotonic() + .6
        while len(reply) < size and time.monotonic() < deadline:
            buf = ctypes.create_string_buffer(size - len(reply)); got = wintypes.DWORD()
            if not self.k.ReadFile(self.handle, buf, len(buf), ctypes.byref(got), None):
                raise ctypes.WinError(ctypes.get_last_error())
            reply.extend(buf.raw[:got.value])
        self.log('IN', hex=reply.hex())
        if len(reply) != size:
            raise RuntimeError('Incomplete serial reply; no automatic retry.')
        return bytes(reply)

    def close(self):
        if self.handle is not None:
            try:
                self.k.SetCommTimeouts(self.handle, ctypes.byref(self.original_timeouts))
            finally:
                self.k.CloseHandle(self.handle); self.handle = None

class Simulator:
    identity = 'SIMULATOR:' + DEVICE_ID
    description = 'Simulation  /  no hardware connection'
    def __init__(self, log, raw=None, device_file=None):
        self.log = log
        self.device_file = Path(device_file) if device_file else None
        if raw is None and self.device_file and self.device_file.exists():
            raw = json.loads(self.device_file.read_text(encoding='utf-8'))
        self.raw = dict(raw or dict(AP=105280, ML=75864, DV=41767, PISTON=-8572))
        self.drill_on = False
        self.pending = {}; self.packets = []; self.closed = False

    def write(self, data):
        self.packets.append(data); self.log('OUT', hex=data.hex())
        if data[:2] == b'\xaf\x11':
            self.drill_on = bool(data[2])
        if data[:2] == b'\xaf\x0f':
            axis = next(a for a in AXES if SELECTORS[a] == data[2])
            self.pending.pop(axis, None)

    def exchange(self, data, size):
        self.write(data)
        if data == bytes((0xaf, 0x12)):
            reply = bytes((0x12, int(self.drill_on))) + bytes(12)
            self.log('IN', hex=reply.hex())
            return reply
        axis = next(a for a in AXES if SELECTORS[a] == data[2])
        if data[1] == 0x0c:
            prior = self.raw[axis]; target = struct.unpack_from('<i', data, 3)[0]
            self.pending[axis] = (target, time.monotonic() + .25)
            reply = b'\x0c' + struct.pack('<ii', prior, target if axis == 'AP' else prior)
        else:
            moving = axis in self.pending
            if moving and time.monotonic() >= self.pending[axis][1]:
                self.raw[axis] = self.pending.pop(axis)[0]; moving = False
            reply = b'\x0e' + struct.pack('<I', int(time.monotonic()*1000) % 2**32)
            reply += bytes((SELECTORS[axis],)) + struct.pack('<iB', self.raw[axis], int(moving))
            reply += bytes(size - len(reply))
        self.log('IN', hex=reply.hex())
        return reply

    def close(self):
        if self.device_file:
            self.device_file.write_text(json.dumps(self.raw), encoding='utf-8')
        self.closed = True
