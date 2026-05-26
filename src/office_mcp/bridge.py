"""The only module that touches the OS.

Everything else builds script strings and runs them here, through the built-in
`osascript`. AppleScript drives actions; JXA (JavaScript) is used for structured
reads because it can `JSON.stringify` a result into a real Python value. Apple
events attach to the already-running app, so edits land in the document the user
has open.
"""

import json
import subprocess


class BridgeError(RuntimeError):
    pass


class NotAuthorized(BridgeError):
    """macOS TCC blocked control of the app (error -1743)."""


def run_applescript(script: str, *args: str) -> str:
    """Run AppleScript. Extra args reach the script's `on run argv` handler,
    which is how user text is passed without quote-escaping."""
    return _run(["osascript", "-", *args], script).strip()


def run_jxa(script: str, *args: str):
    """Run JXA that ends in a `JSON.stringify(...)` expression; return the parsed
    value. Extra args reach the script's `function run(argv)`."""
    out = _run(["osascript", "-l", "JavaScript", "-", *args], script).strip()
    return json.loads(out) if out else None


def _run(cmd: list[str], script: str) -> str:
    proc = subprocess.run(cmd, input=script, capture_output=True, text=True)
    if proc.returncode != 0:
        raise _translate(proc.stderr.strip())
    return proc.stdout


def _translate(stderr: str) -> BridgeError:
    if "-1743" in stderr or "Not authorized" in stderr:
        return NotAuthorized(
            "macOS blocked control of the app. Grant your terminal app Automation "
            "access under System Settings > Privacy & Security > Automation, then retry."
        )
    return BridgeError(stderr or "osascript failed with no error output")
