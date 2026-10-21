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
            if type(self.anchor_backlash[a]) is not int or self.anchor_backlash[a] not in (0, BACKLASH[a]): raise ValueError('Invalid anchor backlash state')
        ref={a:counts[a]-self.anchor_backlash[a] for a in AXES}
        ref['PISTON'] -= 3000 * scales['PISTON']
        return ref


    @classmethod
    def from_current_position(cls, raw_counts, axis_positions_mm, piston_position_nl,
                              current_backlash, counts_per_mm=5225., piston_counts_per_nl=161.36):
        """Derive normalized anchors from measured idle position; never moves hardware."""
        import math
        from .protocol import SIGNS
        if set(raw_counts) != set(AXES) or set(current_backlash) != set(AXES):
            raise ValueError('Supply all four raw counts and verified current backlash states')
        if set(axis_positions_mm) != {'AP','ML','DV'}:
            raise ValueError('Supply measured AP, ML and DV positions in mm')
        values=list(axis_positions_mm.values())+[piston_position_nl,counts_per_mm,piston_counts_per_nl]
        if any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in values):
            raise ValueError('Positions and scales must be finite numbers')
        if counts_per_mm <= 0 or piston_counts_per_nl <= 0 or not 0 <= piston_position_nl <= 5000:
            raise ValueError('Positive scales and Nano 5 uL plunger position 0..5000 nL required')
        for a in AXES:
            if type(raw_counts[a]) is not int or not -(2**31)<=raw_counts[a]<2**31:
                raise ValueError('Invalid raw count')
            if type(current_backlash[a]) is not int or current_backlash[a] not in (0,BACKLASH[a]):
                raise ValueError('Invalid verified current direction history')
        zeros={a:round(raw_counts[a]-current_backlash[a]-SIGNS[a]*counts_per_mm*axis_positions_mm[a])
               for a in ('AP','ML','DV')}
        piston=round(raw_counts['PISTON']-current_backlash['PISTON']+
                     (3000-piston_position_nl)*piston_counts_per_nl)
        result=cls(zeros,piston)
        result.reference(dict(AP=counts_per_mm,ML=counts_per_mm,DV=counts_per_mm,PISTON=piston_counts_per_nl))
        return result
