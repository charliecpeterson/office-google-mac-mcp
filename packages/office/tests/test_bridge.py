"""Bridge mechanics, tested against osascript without needing any Office app."""

import pytest

from office_mcp import bridge


def test_applescript_returns_value():
    assert bridge.run_applescript('return "hi"') == "hi"


def test_applescript_argv_passes_text_unescaped():
    tricky = 'he said "quote" and \\ backslash'
    script = 'on run argv\nreturn item 1 of argv\nend run'
    assert bridge.run_applescript(script, tricky) == tricky


def test_jxa_parses_json():
    assert bridge.run_jxa("JSON.stringify({a: 1, b: [2, 3]})") == {"a": 1, "b": [2, 3]}


def test_error_raises_bridgeerror():
    with pytest.raises(bridge.BridgeError):
        bridge.run_applescript("this is not valid applescript ((")


def test_translate_authorization():
    err = bridge._translate("execution error: Not authorized ... (-1743)")
    assert isinstance(err, bridge.NotAuthorized)
