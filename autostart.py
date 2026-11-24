"""Per-user 'start at Windows sign-in' for the background engine.

The Outlook add-in pages are served by this app on https://localhost:8788. If
the engine is not running, Outlook can only show its generic "Sorry, we can't
load the add-in" message. Starting the engine quietly at sign-in removes the
most common cause of that error. This uses the standard per-user Run key (no
admin rights, no scheduled task, visible in Task Manager > Startup apps).
"""
import sys

RUN_KEY = r'Software\Microsoft\Windows\CurrentVersion\Run'
VALUE = 'DBN Email Studio'


def command(executable=None):
    executable = executable or sys.executable
    return f'"{executable}" --no-browser'


def supported():
    return sys.platform == 'win32' and bool(getattr(sys, 'frozen', False))


def _winreg():
    import winreg  # noqa: PLC0415 - Windows only
    return winreg


def status():
    state = {'supported': supported(), 'enabled': False, 'command': ''}
    if sys.platform != 'win32':
        return state
    winreg = _winreg()
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            value, _ = winreg.QueryValueEx(key, VALUE)
            state.update(enabled=bool(value), command=value)
    except OSError:
        pass
    return state


def set_enabled(enabled):
    if not supported():
        raise OSError('Start at sign-in is available in the installed Windows app only.')
    winreg = _winreg()
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
        if enabled:
            winreg.SetValueEx(key, VALUE, 0, winreg.REG_SZ, command())
        else:
            try:
                winreg.DeleteValue(key, VALUE)
            except FileNotFoundError:
                pass
    return status()
