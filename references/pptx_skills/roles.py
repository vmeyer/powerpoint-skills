"""Semantic roles of placeholders."""

from pptx.enum.shapes import PP_PLACEHOLDER as PH

from .geometry import Box

KICKER_MAX_HEIGHT_CM = 1.5

_ROLES = {
    PH.TITLE: "title",
    PH.CENTER_TITLE: "title",
    PH.VERTICAL_TITLE: "title",
    PH.SUBTITLE: "subtitle",
    PH.BODY: "body",
    PH.OBJECT: "body",
    PH.VERTICAL_BODY: "body",
    PH.VERTICAL_OBJECT: "body",
    PH.PICTURE: "picture",
    PH.BITMAP: "picture",
    PH.TABLE: "table",
    PH.CHART: "chart",
    PH.FOOTER: "footer",
    PH.SLIDE_NUMBER: "number",
    PH.DATE: "date",
}


def placeholder_role(ph_type, box: Box, title_box: Box | None) -> str:
    """Role of a placeholder; a short body/object placeholder above the title is a kicker."""
    role = _ROLES.get(ph_type, "other")
    if (
        role == "body"
        and title_box is not None
        and box.h <= KICKER_MAX_HEIGHT_CM
        and box.bottom <= title_box.y
    ):
        return "kicker"
    return role
