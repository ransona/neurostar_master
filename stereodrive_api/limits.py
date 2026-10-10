"""Configured travel ranges in calibrated mechanical Axis mm / piston nL."""
import math

PISTON_MIN_NL = 500.0
PISTON_MAX_NL = 4500.0
DEFAULT_LIMITS = dict(AP=(0., 40.), ML=(0., 40.), DV=(0., 40.), PISTON=(PISTON_MIN_NL, PISTON_MAX_NL))


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
        if axis == 'PISTON' and not PISTON_MIN_NL <= pair[0] < pair[1] <= PISTON_MAX_NL:
            raise ValueError(f'Piston limits must stay within {PISTON_MIN_NL:g}–{PISTON_MAX_NL:g} nL')
        result[axis] = tuple(float(v) for v in pair)
    return result


def check_target(limits, axis, value):
    low, high = limits[axis]
    if not low - 1e-7 <= value <= high + 1e-7:
        unit = 'nL' if axis == 'PISTON' else 'mm'
        raise ValueError(f'{axis} target {value:g} {unit} exceeds configured travel range {low:g}–{high:g} {unit}')
