"""Unit conversions between EMU, centimetres and guide positions."""

EMU_PER_CM = 360000
EMU_PER_PT = 12700


def cm_to_emu(cm: float) -> int:
    return round(cm * EMU_PER_CM)


def emu_to_cm(emu: int) -> float:
    return round(emu / EMU_PER_CM, 2)


def guide_to_cm(pos: int) -> float:
    """Guide positions are stored in 1/8 pt. Not rounded; callers round for output."""
    return pos / 8 * EMU_PER_PT / EMU_PER_CM
