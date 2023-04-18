import ctypes
import logging

commands = u'/k net stop "camsvc" && net start "camsvc"'
ctypes.windll.shell32.ShellExecuteW(
        None,
        u"runas",
        u"cmd.exe",
        commands,
        None,
        1
    )