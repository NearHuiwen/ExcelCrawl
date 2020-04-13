# -*- coding: utf-8 -*-

#str转bool方法
def str_to_bool(str):
    return True if str.lower() == 'true' else False

import winreg
# 获取windows桌面路径
def get_desktop():
  key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
  return winreg.QueryValueEx(key, "Desktop")[0]