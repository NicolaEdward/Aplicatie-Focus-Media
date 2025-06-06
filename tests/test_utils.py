import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import utils


def test_get_schita_path(tmp_path, monkeypatch):
    folder = tmp_path / "schite"
    folder.mkdir()
    file_path = folder / "X1.png"
    file_path.write_bytes(b"x")
    monkeypatch.setattr(utils, "SCHITE_FOLDER", str(folder))
    assert utils.get_schita_path("X1") == str(file_path)
    file_path.unlink()
    assert utils.get_schita_path("X1") is None
