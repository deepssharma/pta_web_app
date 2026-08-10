"""
Tests for credentials.py. Uses an in-memory fake keyring backend instead
of the real OS credential store, so tests don't touch (or require) an
actual macOS Keychain/Windows Credential Manager/Secret Service.
"""
import keyring
import keyring.backend
import keyring.errors
import pytest

from pta_treasurer import credentials


class FakeKeyring(keyring.backend.KeyringBackend):
    priority = 1

    def __init__(self):
        super().__init__()
        self._store = {}

    def set_password(self, service, username, password):
        self._store[(service, username)] = password

    def get_password(self, service, username):
        return self._store.get((service, username))

    def delete_password(self, service, username):
        key = (service, username)
        if key not in self._store:
            raise keyring.errors.PasswordDeleteError('not found')
        del self._store[key]


class NoBackendKeyring(keyring.backend.KeyringBackend):
    priority = 1

    def set_password(self, service, username, password):
        raise keyring.errors.NoKeyringError('no backend available')

    def get_password(self, service, username):
        raise keyring.errors.NoKeyringError('no backend available')

    def delete_password(self, service, username):
        raise keyring.errors.NoKeyringError('no backend available')


@pytest.fixture
def fake_keyring():
    original = keyring.get_keyring()
    backend = FakeKeyring()
    keyring.set_keyring(backend)
    yield backend
    keyring.set_keyring(original)


@pytest.fixture
def no_backend_keyring():
    original = keyring.get_keyring()
    keyring.set_keyring(NoBackendKeyring())
    yield
    keyring.set_keyring(original)


def test_set_and_get_password_roundtrip(fake_keyring):
    credentials.set_givebacks_password('treasurer@example.org', 's3cret')
    assert credentials.get_givebacks_password('treasurer@example.org') == 's3cret'


def test_get_password_returns_none_when_unset(fake_keyring):
    assert credentials.get_givebacks_password('nobody@example.org') is None


def test_clear_password_removes_it(fake_keyring):
    credentials.set_givebacks_password('treasurer@example.org', 's3cret')
    credentials.clear_givebacks_password('treasurer@example.org')
    assert credentials.get_givebacks_password('treasurer@example.org') is None


def test_clear_password_is_noop_when_nothing_saved(fake_keyring):
    credentials.clear_givebacks_password('nobody@example.org')  # should not raise


def test_different_emails_are_independent(fake_keyring):
    credentials.set_givebacks_password('a@example.org', 'pw-a')
    credentials.set_givebacks_password('b@example.org', 'pw-b')
    assert credentials.get_givebacks_password('a@example.org') == 'pw-a'
    assert credentials.get_givebacks_password('b@example.org') == 'pw-b'


def test_no_backend_raises_clear_error_on_set(no_backend_keyring):
    with pytest.raises(credentials.NoKeyringBackendError):
        credentials.set_givebacks_password('treasurer@example.org', 's3cret')


def test_no_backend_raises_clear_error_on_get(no_backend_keyring):
    with pytest.raises(credentials.NoKeyringBackendError):
        credentials.get_givebacks_password('treasurer@example.org')


def test_no_backend_raises_clear_error_on_clear(no_backend_keyring):
    with pytest.raises(credentials.NoKeyringBackendError):
        credentials.clear_givebacks_password('treasurer@example.org')
