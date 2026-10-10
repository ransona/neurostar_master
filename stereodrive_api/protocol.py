"""Captured StereoDrive serial protocol. No port access in this module."""
import struct
from dataclasses import dataclass

AXES = ('AP', 'ML', 'DV', 'PISTON')
SELECTORS = dict(AP=0x40, ML=0x50, DV=0x60, PISTON=0x70)
SIGNS = dict(AP=-1, ML=1, DV=-1, PISTON=1)  # Increasing mechanical mm versus raw counts.
BACKLASH = dict(AP=522, ML=261, DV=52, PISTON=5825)
STEPS = (.01, .02, .05, .1, .2, .5, 1.)

@dataclass(frozen=True)
class Status:
    raw: int
    moving: bool
    clock: int

def poll(axis):
    return bytes((0xaf, 0x0e, SELECTORS[axis]))

def stop(axis):
    return bytes((0xaf, 0x0f, SELECTORS[axis]))

def target(axis, raw, direction, speed_mm_s=2):
    if type(raw) is not int or not -(2**31) <= raw < 2**31:
        raise ValueError('Target must be a signed 32-bit motor count.')
    if direction not in (-1, 1):
        raise ValueError('Direction must be -1 or +1.')
    if speed_mm_s not in (1, 2): raise ValueError('Validated axis profiles are 1 or 2 mm/s')
    flag = int(direction > 0) if axis == 'DV' else 1
    return bytes((0xaf, 0x0c, SELECTORS[axis])) + struct.pack(
        '<iHHHBBBB', raw, 1820 if axis == 'PISTON' else 653 * int(speed_mm_s), 8000, 8000,
        0, 1 if axis == 'PISTON' or speed_mm_s == 2 else 2, 0 if axis == 'PISTON' else flag, 4)

def decode_status(axis, data):
    size = 21 if axis == 'AP' else 26
    if len(data) != size or data[0] != 0x0e or data[5] != SELECTORS[axis]:
        raise RuntimeError('Unexpected status reply; no guessed resynchronization.')
    if data[10] not in (0, 1):
        raise RuntimeError('Unknown motion status.')
    return Status(struct.unpack_from('<i', data, 6)[0], bool(data[10]),
                  struct.unpack_from('<I', data, 1)[0])

def validate_ack(axis, data, prior, requested):
    # AP echoes prior/target. ML and DV echo prior/prior in completed captures.
    expected = (prior, requested if axis == 'AP' else prior)
    if len(data) != 9 or data[0] != 0x0c or struct.unpack_from('<ii', data, 1) != expected:
        raise RuntimeError('Unexpected target acknowledgement; command will not be retried.')


def drill_power(enabled):
    if type(enabled) is not bool: raise ValueError('Drill power requires bool')
    return bytes((0xaf, 0x11, int(enabled)))

def drill_query():
    return bytes((0xaf, 0x12))

def decode_drill(data):
    if len(data) != 14 or data[0] != 0x12 or data[1] not in (0,1):
        raise RuntimeError('Unexpected drill-state reply')
    return bool(data[1])
