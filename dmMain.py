from win32com.client import Dispatch
import os
from random import uniform
from time import sleep


class Operation:

    def __init__(self, dm, hwnd):
        self.dm = dm
        self.hwnd = hwnd
        # self.dm.Reg('vacation6c113ef949e77e259bfadb84959f6cbc', '')
        self.dm.Ver()
        self.bind()

    def bind(self):
        # self.dm.BindWindowEx(self.hwnd, "normal", "normal", "normal", "", 0)
        self.dm.BindWindowEx(self.hwnd, "gdi", "windows", "windows", "", 0)
        self.dm.SetSimMode(0)
        self.dm.EnableRealKeypad(1)
        self.dm.EnableRealMouse(2, 20, 30)
        self.dm.SetKeypadDelay("normal", 70)
        self.dm.SetClientSize(self.hwnd, 596, 446)
        print(self.dm.GetClientSize(self.hwnd))
        print('绑定成功')

    def getwindowsize(self):
        ret = self.dm.GetClientSize(self.hwnd)
        self.width, self.height = ret[1], ret[2]
        print(self.width, self.height)

    def leftclick(self, xf, yf, ran_x, ran_y, delay=uniform(0.3, 0.5)):
        x = xf + uniform(0, ran_x)
        y = yf + uniform(0, ran_y)
        self.dm.MoveTo(x, y)
        self.dm.LeftClick()
        sleep(delay)

    def keypress(self, n):
        self.dm.KeyPress(n)

    def keyup(self, n):
        self.dm.KeyUp(n)

    def keydown(self, n):
        self.dm.KeyDown(n)


def regsvr():
    try:
        dm_1 = Dispatch('dm.dmsoft')
    except Exception:
        os.system(r'regsvr32 /s %s\dm.dll' % os.getcwd())
        dm_1 = Dispatch('dm.dmsoft')
    print(dm_1.Ver())
    return dm_1