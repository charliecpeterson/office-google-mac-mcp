"""The only module that touches the OS.

Everything else builds script strings and runs them here, through the built-in
`osascript`. AppleScript drives actions; JXA (JavaScript) is used for structured
reads because it can `JSON.stringify` a result into a real Python value. Apple
events attach to the already-running app, so edits land in the document the user
has open. `screenshot` is the other OS touch point — it captures an app window so
the model can see its own work.
"""

import json
import os
import subprocess
import tempfile
from pathlib import Path


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


def screenshot(app_name: str) -> bytes:
    """Capture the app's largest on-screen window as PNG bytes, even when occluded.

    The window id comes from Quartz so we don't have to raise or move the window,
    then the built-in `screencapture` grabs just that window. Needs Screen
    Recording permission, which is separate from the Automation permission.
    """
    window_id = _window_id(app_name)
    if window_id is None:
        raise BridgeError(f"no on-screen window found for {app_name!r}; is it open?")
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        proc = subprocess.run(
            ["screencapture", f"-l{window_id}", "-o", "-x", path],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            err = proc.stderr.strip()
            if "could not create image" in err.lower():
                raise NotAuthorized(
                    "macOS blocked the screen capture. Grant your terminal app Screen "
                    "Recording access under System Settings > Privacy & Security > Screen "
                    "Recording, then restart the terminal."
                )
            raise BridgeError(err or "screencapture failed")
        return Path(path).read_bytes()
    finally:
        os.unlink(path)


def _window_id(app_name: str) -> int | None:
    import Quartz

    best = None
    windows = Quartz.CGWindowListCopyWindowInfo(
        Quartz.kCGWindowListOptionOnScreenOnly | Quartz.kCGWindowListExcludeDesktopElements,
        Quartz.kCGNullWindowID,
    )
    for w in windows:
        if w.get("kCGWindowOwnerName") != app_name:
            continue
        bounds = w["kCGWindowBounds"]
        area = bounds["Width"] * bounds["Height"]
        if best is None or area > best[0]:
            best = (area, int(w["kCGWindowNumber"]))
    return best[1] if best else None
