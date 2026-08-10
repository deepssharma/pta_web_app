"""
credentials.py
Givebacks password storage via the OS-native credential store (macOS
Keychain / Windows Credential Manager / Linux Secret Service), through
the `keyring` library. This is the only place a Givebacks password ever
passes through -- it is never written to org_config.json or any other
file the app controls, never logged, never printed.

The AI Assistant has no credentials module of its own -- it runs
against a local Ollama server, which needs no API key/secret at all.

Non-secret Givebacks identifiers (org URL, email, cause ID) live in
OrgConfig/org_config.json alongside everything else; only the password
is handled here.
"""

import keyring
import keyring.errors

GIVEBACKS_SERVICE_NAME = 'pta-treasurer-givebacks'


class NoKeyringBackendError(Exception):
    """Raised when no OS credential store is available on this machine."""


def set_givebacks_password(email: str, password: str) -> None:
    try:
        keyring.set_password(GIVEBACKS_SERVICE_NAME, email, password)
    except keyring.errors.NoKeyringError as e:
        raise NoKeyringBackendError(
            'No OS credential store is available on this machine -- '
            'Givebacks auto-download needs one to store the password '
            'securely. Use manual CSV export instead.'
        ) from e


def get_givebacks_password(email: str) -> str | None:
    try:
        return keyring.get_password(GIVEBACKS_SERVICE_NAME, email)
    except keyring.errors.NoKeyringError as e:
        raise NoKeyringBackendError(
            'No OS credential store is available on this machine -- '
            'Givebacks auto-download needs one to retrieve the saved '
            'password. Use manual CSV export instead.'
        ) from e


def clear_givebacks_password(email: str) -> None:
    try:
        keyring.delete_password(GIVEBACKS_SERVICE_NAME, email)
    except keyring.errors.PasswordDeleteError:
        pass  # nothing saved for this email -- nothing to do
    except keyring.errors.NoKeyringError as e:
        raise NoKeyringBackendError(
            'No OS credential store is available on this machine.'
        ) from e
