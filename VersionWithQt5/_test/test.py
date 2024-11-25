import webbrowser
import winreg
#浏览器注册表信息

_browser_regs = {
    "IE":r"SOFTWARE\\Clients\\StartMenuInternet\\IEXPLORE.EXE\\DefaultIcon",
    "chrome":r"SOFTWARE\\Clients\\StartMenuInternet\\Google Chrome\\DefaultIcon",
    "edge":r"SOFTWARE\\Clients\\StartMenuInternet\\Microsoft Edge\\DefaultIcon",
    "firefox":r"SOFTWARE\\Clients\\StartMenuInternet\\FIREFOX.EXE\\DefaultIcon",
    "360":r"SOFTWARE\\Clients\\StartMenuInternet\\360Chrome\\DefaultIcon",
}
def get_browser_path(browser):
    try:
        key=winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,_browser_regs.get(browser))
    except FileNotFoundError:
        return None
    value,_type=winreg.QueryValueEx(key,"")
    return value.split(",")[0]

print(get_browser_path("edge"))
