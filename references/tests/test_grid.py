import json

import pytest

from pptx_skills.cli import main
from pptx_skills.errors import PptxSkillsError
from pptx_skills.geometry import Box
from pptx_skills.grid import grid_boxes


def test_grid_four_columns():
    rows = grid_boxes(Box(1.5, 4.82, 30.87, 11.88), cols=4, gutter=0.3)
    assert len(rows) == 1 and len(rows[0]) == 4
    assert rows[0][0].x == 1.5 and rows[0][3].right == pytest.approx(32.37, abs=0.001)
    assert rows[0][1].x - rows[0][0].right == pytest.approx(0.3, abs=0.01)


def test_grid_rows_and_row_gutter():
    rows = grid_boxes(Box(0, 0, 10, 10), cols=2, rows=3, gutter=1, row_gutter=0.5)
    assert len(rows) == 3
    assert rows[2][1].right == 10 and rows[2][1].bottom == 10
    assert rows[1][0].y - rows[0][0].bottom == pytest.approx(0.5, abs=0.01)


@pytest.mark.parametrize("kwargs", [{"cols": 0}, {"cols": 1, "rows": 0}, {"cols": 2, "gutter": -1}, {"cols": 3, "gutter": 5}])
def test_grid_rejects_bad_input(kwargs):
    with pytest.raises(PptxSkillsError):
        grid_boxes(Box(0, 0, 10, 10), **kwargs)


def test_cli_json_and_text(deck_dir_from_default, capsys):
    assert main(["grid", str(deck_dir_from_default), "--cols", "2", "--rows", "2", "--json"]) == 0
    data = json.loads(capsys.readouterr().out)
    assert len(data["rows"]) == 2 and len(data["rows"][0]) == 2 and len(data["area"]) == 4
    assert main(["grid", str(deck_dir_from_default), "--cols", "2", "--area", "1,2,10,5"]) == 0
    out = capsys.readouterr().out
    assert "add_shape" in out and "1,2 " in out


def test_cli_full_area_and_errors(deck_dir_from_default, capsys):
    assert main(["grid", str(deck_dir_from_default), "--cols", "1", "--area", "full", "--json"]) == 0
    area = json.loads(capsys.readouterr().out)["area"]
    assert area[:2] == [0, 0]
    assert main(["grid", str(deck_dir_from_default), "--cols", "1", "--area", "1,2,3"]) == 2
    assert main(["grid", str(deck_dir_from_default), "--cols", "0"]) == 2
