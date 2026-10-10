"""Configured travel ranges in calibrated mechanical Axis mm / piston nL."""
import math

DEFAULT_LIMITS = dict(AP=(0., 40.), ML=(0., 40.), DV=(0., 40.), PISTON=(0., 5000.))


def validate_limits(limits=None):
    limits = DEFAULT_LIMITS if limits is None else limits
    if not isinstance(limits, dict) or set(limits) != set(DEFAULT_LIMITS):
        raise ValueError('Supply AP, ML, DV and PISTON travel ranges')
    result = {}
    for axis, pair in limits.items():
        if (not isinstance(pair, (tuple, list)) or len(pair) != 2
                or any(isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) for v in pair)
                or pair[0] >= pair[1]):
            raise ValueError(f'{axis}: finite minimum must be less than maximum')
        if axis == 'PISTON' and not 0 <= pair[0] < pair[1] <= 5000:
            raise ValueError('Nano 5 µL piston limits must stay within 0–5000 nL')
        result[axis] = tuple(float(v) for v in pair)
    return result


def check_target(limits, axis, value):
    low, high = limits[axis]
    if not low - 1e-7 <= value <= high + 1e-7:
        unit = 'nL' if axis == 'PISTON' else 'mm'
        raise ValueError(f'{axis} target {value:g} {unit} exceeds configured travel range {low:g}–{high:g} {unit}')
