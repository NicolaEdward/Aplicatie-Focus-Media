import os
import sqlite3
import importlib

import db


def setup_module(module):
    # Use an in-memory SQLite database for testing
    test_conn = db._ConnWrapper(sqlite3.connect(':memory:'), False)
    db.conn = test_conn
    db.cursor = test_conn.cursor()
    db.init_db()
    db.refresh_location_cache()


def test_parse_port():
    assert db._parse_port('3306') == 3306
    assert db._parse_port('') is None
    assert db._parse_port(None) is None
    assert db._parse_port(' 123 ') == 123

    try:
        db._parse_port('abc')
    except ValueError:
        pass
    else:
        assert False, 'Expected ValueError'


def test_hash_and_verify_password():
    pw = 'secret'
    hashed = db._hash_password(pw)
    assert hashed != pw
    assert db._verify_password(hashed, pw)
    assert not db._verify_password(hashed, 'other')


def test_create_and_get_user():
    username = 'testuser'
    db.create_user(username, 'pw123', role='seller')
    user = db.get_user(username)
    assert user['username'] == username
    assert db.check_login(username, 'pw123')
    assert not db.check_login(username, 'wrong')


