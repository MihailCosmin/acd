"""The last line of defence when a reference cannot be resolved: ask.

A checker resolves a schema or a BREX from the checked object's own folder,
then from whatever the caller registered, then from the `res/` tree found by
`resources`. When all of that comes up empty the run is about to report a
reference it never actually checked -- so rather than silently degrade, the
checker asks once for a folder to look in.

Asking from inside a library is only safe if it can never block something
that has no one to answer. Every guard here exists for that:

- `prompting_enabled()` is False under pytest, under `ACD_NO_PROMPT`, and
  whenever no display is available, so a batch job, a CI run or a service
  degrades to reporting the reference as unresolved instead of hanging.
- A prompt is only ever raised once per checker run, listing everything that
  is missing together, so a folder of 90 objects sharing one missing schema
  asks once rather than ninety times.
- Qt is used when the calling application is already running a Qt event loop
  (the ALTHOM applications are PySide6), because opening a tkinter window
  underneath a live QApplication is a good way to deadlock. tkinter is the
  fallback for a plain script, and "no toolkit" is a perfectly good answer.

Nothing here raises. Every failure path returns None, which the caller reads
as "still unresolved".
"""

from os import environ
from sys import modules
from sys import platform

# Set to any non-empty value to suppress every prompt: the checkers then
# report an unresolvable reference rather than asking for it. Intended for
# batch jobs, services and CI.
NO_PROMPT_ENV = "ACD_NO_PROMPT"

# Module-level off switch, for a caller that wants library-wide silence
# without touching the environment. `set_prompting(False)` is the programmatic
# equivalent of `ACD_NO_PROMPT`.
_ENABLED = True


def set_prompting(enabled: bool = True) -> None:
    """Turn the missing-resource prompts on or off process-wide.

    Args:
        enabled (bool): whether a checker may ask for a missing location
    """
    global _ENABLED  # pylint: disable=global-statement
    _ENABLED = enabled


def _has_display() -> bool:
    """Whether there is anything to draw a dialog on.

    Windows and macOS always have a window server available to a desktop
    process; X11/Wayland does not, and a headless Linux box is exactly the
    case that must never block.
    """
    if platform.startswith("win") or platform == "darwin":
        return True
    return bool(environ.get("DISPLAY") or environ.get("WAYLAND_DISPLAY"))


def prompting_enabled() -> bool:
    """Whether a prompt may be raised right now.

    Returns:
        bool: False under pytest, under `ACD_NO_PROMPT`, after
            `set_prompting(False)`, or with no display available
    """
    if not _ENABLED:
        return False
    if environ.get(NO_PROMPT_ENV):
        return False
    # Never interrupt a test run: a modal dialog in a suite hangs it until it
    # is killed, and the failure looks like a timeout rather than a prompt.
    if "pytest" in modules or environ.get("PYTEST_CURRENT_TEST"):
        return False
    return _has_display()


def _qt_application():
    """The running QApplication, if the caller has one.

    Returns the instance rather than creating one: a library must not own the
    application object, and a Qt dialog raised without an event loop of its
    own belongs to whatever loop is already running.
    """
    qtwidgets = modules.get("PySide6.QtWidgets") or modules.get("PySide2.QtWidgets")
    if qtwidgets is None:
        try:
            from PySide6 import QtWidgets as qtwidgets  # noqa: N813
        except ImportError:
            try:
                from PySide2 import QtWidgets as qtwidgets  # noqa: N813
            except ImportError:
                return None, None
    return qtwidgets, qtwidgets.QApplication.instance()


def _ask_qt(title: str, message: str) -> str:
    """Qt folder picker, preceded by a message naming what is missing."""
    qtwidgets, application = _qt_application()
    if qtwidgets is None:
        return None
    owns_application = False
    if application is None:
        # A plain script that happens to have PySide6 installed: stand up a
        # throwaway application so the dialog has a loop, and tear it down.
        application = qtwidgets.QApplication([])
        owns_application = True
    try:
        qtwidgets.QMessageBox.information(
            None, title, message, qtwidgets.QMessageBox.Ok
        )
        return qtwidgets.QFileDialog.getExistingDirectory(
            None, f"{title} - select the folder that contains them"
        ) or None
    except Exception:  # pylint: disable=broad-except
        return None
    finally:
        if owns_application:
            application.quit()


def _ask_tk(title: str, message: str) -> str:
    """tkinter folder picker, for a script with no Qt around."""
    try:
        from tkinter import Tk
        from tkinter import messagebox
        from tkinter import filedialog
    except ImportError:
        return None
    root = None
    try:
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(title, message, parent=root)
        return filedialog.askdirectory(
            parent=root, title=f"{title} - select the folder that contains them"
        ) or None
    except Exception:  # pylint: disable=broad-except
        return None
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:  # pylint: disable=broad-except
                pass


def ask_for_folder(what: str, missing: list, limit: int = 12) -> str:
    """Ask once for a folder holding the references that could not be found.

    Everything missing is listed in a single dialog rather than one dialog per
    reference: a folder run where ninety objects share one unresolvable schema
    has one thing wrong with it, not ninety.

    Args:
        what (str): what is missing, in plural and lower case -- "schemas" or
            "BREX data modules"
        missing (list): the references that could not be resolved; duplicates
            and order are the caller's business
        limit (int): how many to name before summarising the remainder, so a
            pathological run cannot produce a dialog taller than the screen

    Returns:
        str: the chosen directory, or None if prompting is unavailable, the
            dialog could not be raised, or the user cancelled
    """
    if not missing or not prompting_enabled():
        return None

    listed = list(missing)
    shown = "\n".join(f"    {_}" for _ in listed[:limit])
    if len(listed) > limit:
        shown += f"\n    ... and {len(listed) - limit} more"

    title = f"Missing {what}"
    message = (
        f"{len(listed)} {what} could not be found in the folder being checked, "
        f"in any registered search path, or in the bundled resources:\n\n{shown}\n\n"
        "Select the folder that contains them, or cancel to report them as "
        "unresolved."
    )

    qtwidgets, application = _qt_application()
    if qtwidgets is not None and application is not None:
        # The caller is already running Qt; anything else risks a deadlock.
        return _ask_qt(title, message)
    return _ask_tk(title, message) or _ask_qt(title, message)
