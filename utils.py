import socket
import sys
import win32api
import win32con


def find_path(name="msedge.exe") -> str:
    path = ""
    try:
        key = rf'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{name}'
        #通过获取Windows注册表查找软件
        key = win32api.RegOpenKey(win32con.HKEY_LOCAL_MACHINE, key, 0, win32con.KEY_READ)
        info2 = win32api.RegQueryInfoKey(key)
        for j in range(0, info2[1]):
            key_value = win32api.RegEnumValue(key, j)[1]
            if key_value.upper().endswith(name.upper()):
                path = key_value
                break
        win32api.RegCloseKey(key)
    except Exception as e:
        print(e)
        pass
    return path	 #返回查找到的安装路径


def is_port_used() -> bool:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect(("127.0.0.1", 9222))
        s.shutdown(2)
        print("9222端口已被占用")
        return True
    except:
        print("9222端口未被占用")
        return False


def ez_get_object(j, list_keys):
    """
    快速获取json转换后的对象子元素
    j = {
        'a':
            {'b':{
                b1: {'name':'xiao'},
                b2: {'name':'2b'}
            }}
    }

    >>> ez_get_object(j, 'a,b,b2')
    {'name':'2b'}

    j: json转换后的对象
    list_keys： 逗号分隔的字符串
    """
    ret = j
    for k in list_keys.split(','):
        ret = ret.get(k.strip())

    return ret
