
import win32api, win32con
import time

while 1:
    win32api.SetCursorPos((725,912))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,725,912,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,725,912,0,0)
    time.sleep(17)