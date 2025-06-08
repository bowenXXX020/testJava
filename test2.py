import ctypes
import os

from comtypes.client import CreateObject

patch = ctypes.windll.LoadLibrary(os.path.dirname(__file__) + './RegDll.dll')
patch.SetDllPathW(os.path.dirname(__file__) + './dm.dll', 0)
dm = CreateObject('dm.dmsoft')  # 创建对象
print('免注册调用初始化成功 版本号为:', dm.ver())