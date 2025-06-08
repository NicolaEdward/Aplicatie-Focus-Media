import UI.utils as ui_utils


def test_parse_port_ok():
    assert ui_utils._parse_port('3306') == 3306
    assert ui_utils._parse_port(' 443 ') == 443
    assert ui_utils._parse_port(None) is None
    assert ui_utils._parse_port('') is None


def test_parse_port_invalid():
    import pytest
    with pytest.raises(ValueError):
        ui_utils._parse_port('abc')


def test_get_db_connection_params(monkeypatch):
    monkeypatch.setenv('MYSQL_HOST', 'localhost:3307')
    monkeypatch.setenv('MYSQL_USER', 'u')
    monkeypatch.setenv('MYSQL_PASSWORD', 'p')
    monkeypatch.setenv('MYSQL_DATABASE', 'd')

    class DummyConn:
        def close(self):
            pass

    def dummy_connect(**kwargs):
        return DummyConn()

    monkeypatch.setattr(ui_utils.mysql.connector, 'connect', dummy_connect)
    conn = ui_utils.get_db_connection()
    assert isinstance(conn, DummyConn)
    conn.close()
