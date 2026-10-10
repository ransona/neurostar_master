"""Count anchors for absolute axis mm and estimated piston nL."""
from dataclasses import dataclass, field
from .protocol import AXES, BACKLASH

@dataclass(frozen=True)
class Calibration:
    axis_zero_counts: dict
    piston_3000_count: int
    anchor_backlash: dict = field(default_factory=lambda: dict.fromkeys(AXES, 0))

    def reference(self, scales):
        if set(self.axis_zero_counts) != {'AP','ML','DV'}:
            raise ValueError('Supply AP, ML and DV zero counts')
        counts = dict(self.axis_zero_counts, PISTON=self.piston_3000_count)
        if set(self.anchor_backlash) != set(AXES): raise ValueError('Supply all anchor backlash states')
        for a in AXES:
            if type(counts[a]) is not int or not -(2**31) <= counts[a] < 2**31:
                raise ValueError('Anchors must be signed 32-bit integer counts')
            if self.anchor_backlash[a] not in (0, BACKLASH[a]): raise ValueError('Invalid anchor backlash state')
        ref={a:counts[a]-self.anchor_backlash[a] for a in AXES}
        ref['PISTON'] -= 3000 * scales['PISTON']
        return ref
