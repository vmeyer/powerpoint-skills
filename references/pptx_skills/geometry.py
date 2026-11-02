"""Axis-aligned boxes in centimetres."""

from dataclasses import dataclass


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
