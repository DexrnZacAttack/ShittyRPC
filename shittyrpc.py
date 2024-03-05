from pypresence import Presence
import time
import mouse
from datetime import datetime, timedelta
import platform
import ctypes
import distro
import pygetwindow
if platform.system().lower().startswith('lin'):
    import dbus
if platform.system().lower().startswith('win'):
    import win32gui, win32process, psutil

lastMPos = None
lastCurMoveTime = None
discordRPC = None

# Change this to false if you don't want it to show what program you're currently active in.
# If this is false, it will simply say "Using the computer" instead.
showActivePG = True
# Change this to true if you want to display the Window title (note this probably will leak information)
showWinTitle = False


def getPlatform():
    if platform.system().lower().startswith('lin'):
        return distro.name() + " " + distro.version() + " (" + distro.codename().title() + ")"
    elif platform.system().lower().startswith('win'):
        return "Windows "  + platform.release() + " build " + platform.version()


def getActiveWindow(type = 0):
    if platform.system().lower().startswith('win'):
        if showWinTitle is not True or type == 2:
            # https://stackoverflow.com/questions/65362756/getting-process-name-of-focused-window-with-python
            try:
                pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
                return(psutil.Process(pid[-1]).name())
            except:
                pass
        elif showWinTitle is True:
            # THIS LEAKS THE ENTIRE WINDOW TITLE
            return pygetwindow.getActiveWindow()

def isOnLockScreen():
    if platform.system().lower().startswith('dar'):
        pass
    if platform.system().lower().startswith('win'):
        if ctypes.windll.user32.GetForegroundWindow() == 0:
            return True
        elif getActiveWindow(2) == "LockApp.exe":
            # thank you Windows
            return True
        else:
            return False
    else:
        session_bus = dbus.SessionBus()
        screensaver_proxy = session_bus.get_object('org.gnome.ScreenSaver', '/org/gnome/ScreenSaver')
        interface = dbus.Interface(screensaver_proxy, 'org.gnome.ScreenSaver')
        return interface.GetActive()

# Dexrn: This only checks if the cursor has moved, this can probably be improved
def getCursorStats(type):
    global lastMPos
    global lastCurMoveTime
    # if I ever come back to this project I might fix this area up
    if type == 1:
        # return mouse pos, prob won't help
        return mouse.get_position()
    elif type == 2:
        # return true or false if the mouse hasnt moved since the last check
        mPos = mouse.get_position()
        if lastMPos != mPos:
            return True
        else:
            return False
        lastMPos = mPos
    elif type == 3:
        # return a string if the mouse hasnt moved for a set amount of time
        print("3")
        mPos = mouse.get_position()
        if lastMPos != mPos:
            lastCurMoveTime = 0
            lastMPos = mPos
            if showActivePG is True:
                return "Active window: " + getActiveWindow()
            else:
                return "Using the computer"
        else:
            lastMPos = mPos
            lastCurMoveTime += 1
            if lastCurMoveTime is not None:
                if lastCurMoveTime > 3600 and isOnLockScreen() is not True:
                    return "Probably eeping (No cursor movement for: " + str(timedelta(seconds=lastCurMoveTime)) + ")"
                elif lastCurMoveTime > 5 and isOnLockScreen() is not True:
                    return "Away (No cursor movement for: " + str(timedelta(seconds=lastCurMoveTime)) + ")"
                else:
                    if showActivePG is True:
                        return "Active window: " + getActiveWindow()
                    else:
                        return "Using the computer"
            else:
                if showActivePG is True:
                    return "Active window: " + getActiveWindow()
                else:
                    return "Using the computer"
    elif type == 4:
        return "On lock screen"

# https://stackoverflow.com/questions/68712041/how-to-get-system-boot-time-with-millisecond-precision
# Dexrn: using this because it doesn't need to be that accurate
# Modified a little to work cross platform
def uptime_delta():
    """
    Get delta between system uptime and unix timestamp
    :return: float of unix timestamp delta, result of unixtimestamp now - uptime
    """
    if platform.system().lower().startswith('win'):
        kernel32 = ctypes.windll.kernel32
        uptime = kernel32.GetTickCount64() / 1000.0
        now = datetime.now().timestamp()
        return now - uptime
    elif not platform.system().lower().startswith('dar'):
        with open("/proc/uptime") as f:
            uptime = float(f.read().split()[0])
            now = datetime.now().timestamp()
            return now - uptime

def getStatus(discordRPC):
    print("Refreshed")
    if isOnLockScreen():
            discordRPC.update(
            # Dexrn: There might be a way to not calculate this over and over again.
            details=getPlatform(),
            start=bootTime,
            state=getCursorStats(4))
    elif not isOnLockScreen():
            discordRPC.update(
            details=getPlatform(),
            start=bootTime,
            state=getCursorStats(3))


def connect():
    print("Connecting")
    try:
        discordRPC = Presence("1214151460288462909")
        discordRPC.connect()
        if True:
            print("Connected")
        while True:
            getStatus(discordRPC)
            time.sleep(1)
    except Exception as e:
        print("erorr")
        print(str(e))
        if discordRPC != None:
            discordRPC.close()
        connect()

bootTime = uptime_delta()
connect()
