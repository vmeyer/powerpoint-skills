"""Axis-aligned boxes in centimetres."""

from dataclasses import dataclass

from pptx.oxml.ns import qn


@dataclass(frozen=True)
class Box:
    x: float
    y: float
    w: float
    h: float

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    def contains(self, other: "Box", tol: float = 0.01) -> bool:
        return (
            other.x >= self.x - tol
            and other.y >= self.y - tol
            and other.right <= self.right + tol
            and other.bottom <= self.bottom + tol
        )

    def intersects(self, other: "Box") -> float:
        """Overlap area in cm², 0.0 when the boxes do not overlap."""
        dx = min(self.right, other.right) - max(self.x, other.x)
        dy = min(self.bottom, other.bottom) - max(self.y, other.y)
        return dx * dy if dx > 0 and dy > 0 else 0.0

    def as_list(self) -> list[float]:
        return [self.x, self.y, self.w, self.h]


def _xfrm(element):
    found = element.xpath("./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm | ./p:xfrm")
    return found[0] if found else None


def absolute_box(shape) -> Box:
    """Slide-absolute box of a shape in cm, resolving every enclosing group transform.

    Placeholders without their own ``xfrm`` use the position python-pptx inherits from the layout.
    """
    from .units import EMU_PER_CM

    x, y = float(shape.left or 0), float(shape.top or 0)
    w, h = float(shape.width or 0), float(shape.height or 0)
    parent = shape._element.getparent()
    while parent is not None and parent.tag.endswith("}grpSp"):
        xfrm = _xfrm(parent)
        if xfrm is not None:
            off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
            ch_off, ch_ext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
            if None not in (off, ext, ch_off, ch_ext):
                sx = int(ext.get("cx")) / int(ch_ext.get("cx")) if int(ch_ext.get("cx")) else 1.0
                sy = int(ext.get("cy")) / int(ch_ext.get("cy")) if int(ch_ext.get("cy")) else 1.0
                x = int(off.get("x")) + (x - int(ch_off.get("x"))) * sx
                y = int(off.get("y")) + (y - int(ch_off.get("y"))) * sy
                w, h = w * sx, h * sy
        parent = parent.getparent()
    return Box(*(round(v / EMU_PER_CM, 2) for v in (x, y, w, h)))


def _snap_edge(value: float, lines: list[float], tol: float) -> float:
    near = min(lines, key=lambda line: abs(line - value), default=None)
    return near if near is not None and abs(near - value) <= tol else value


def snap_box(box: Box, xs: list[float], ys: list[float], tol: float) -> Box:
    """Snap each edge to the nearest vertical (`xs`) or horizontal (`ys`) line within `tol` cm.

    Edges move independently, so a snapped left edge with a free right edge changes the width.
    """
    left, right = _snap_edge(box.x, xs, tol), _snap_edge(box.right, xs, tol)
    top, bottom = _snap_edge(box.y, ys, tol), _snap_edge(box.bottom, ys, tol)
    return Box(round(left, 2), round(top, 2), round(right - left, 2), round(bottom - top, 2))
