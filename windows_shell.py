"""Brand only DBN Email Studio's own Edge app windows; preserve their contents."""
import ctypes
from ctypes import wintypes
from pathlib import Path
import sys
import threading
import time
import uuid

_icons=[]
_worker=None
_lock=threading.Lock()


class GUID(ctypes.Structure):
    _fields_=[('data',ctypes.c_byte*16)]


class PROPERTYKEY(ctypes.Structure):
    _fields_=[('fmtid',GUID),('pid',wintypes.DWORD)]


class VALUE(ctypes.Union):
    _fields_=[('pointer',ctypes.c_void_p),('alignment',ctypes.c_longlong),('bytes',ctypes.c_byte*16)]


class PROPVARIANT(ctypes.Structure):
    _fields_=[('vt',ctypes.c_ushort),('reserved',ctypes.c_ushort*3),('value',VALUE)]


def guid(value):return GUID.from_buffer_copy(uuid.UUID(value).bytes_le)


def _properties(hwnd, icon_path):
    interface=ctypes.c_void_p()
    iid=guid('886d8eeb-8cf2-4446-8d02-cdba1dbdcf99')
    get=ctypes.windll.shell32.SHGetPropertyStoreForWindow
    get.argtypes=[wintypes.HWND,ctypes.POINTER(GUID),ctypes.POINTER(ctypes.c_void_p)]
    get.restype=ctypes.c_long
    if get(hwnd,ctypes.byref(iid),ctypes.byref(interface))!=0:return
    vtable=ctypes.cast(interface,ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    setter=ctypes.WINFUNCTYPE(ctypes.c_long,ctypes.c_void_p,ctypes.POINTER(PROPERTYKEY),ctypes.POINTER(PROPVARIANT))(vtable[6])
    commit=ctypes.WINFUNCTYPE(ctypes.c_long,ctypes.c_void_p)(vtable[7])
    release=ctypes.WINFUNCTYPE(ctypes.c_ulong,ctypes.c_void_p)(vtable[2])
    init=ctypes.windll.propsys.InitPropVariantFromString
    init.argtypes=[wintypes.LPCWSTR,ctypes.POINTER(PROPVARIANT)];init.restype=ctypes.c_long
    clear=ctypes.windll.ole32.PropVariantClear;clear.argtypes=[ctypes.POINTER(PROPVARIANT)]
    values={5:'DBN.EmailStudio.Blue',3:str(icon_path)+',0',4:'DBN Email Studio'}
    if getattr(sys,'frozen',False):values[2]='"'+sys.executable+'"'
    try:
        for pid,text in values.items():
            key=PROPERTYKEY(guid('9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3'),pid)
            value=PROPVARIANT()
            if init(text,ctypes.byref(value))==0:
                try:setter(interface,ctypes.byref(key),ctypes.byref(value))
                finally:clear(ctypes.byref(value))
        commit(interface)
    finally:release(interface)


def refresh(root, data_dir):
    global _worker
    if sys.platform!='win32':return
    with _lock:
        if _worker and _worker.is_alive():return
        def work():
            try:
                ctypes.windll.ole32.CoInitializeEx(None,0)
                icon_path=Path(data_dir)/'DBN Email Studio Blue.ico'
                icon_path.parent.mkdir(parents=True,exist_ok=True)
                icon_path.write_bytes((Path(root)/'web'/'icons'/'email-blue-v2.ico').read_bytes())
                user=ctypes.windll.user32
                load=user.LoadImageW;load.argtypes=[wintypes.HINSTANCE,wintypes.LPCWSTR,wintypes.UINT,ctypes.c_int,ctypes.c_int,wintypes.UINT];load.restype=wintypes.HANDLE
                if not _icons:
                    _icons.extend([load(None,str(icon_path),1,size,size,16) for size in (16,32)])
                send=user.SendMessageTimeoutW
                send.argtypes=[wintypes.HWND,wintypes.UINT,ctypes.c_size_t,ctypes.c_ssize_t,wintypes.UINT,wintypes.UINT,ctypes.POINTER(ctypes.c_size_t)]
                callback_type=ctypes.WINFUNCTYPE(wintypes.BOOL,wintypes.HWND,wintypes.LPARAM)
                user.GetWindowTextW.argtypes=[wintypes.HWND,wintypes.LPWSTR,ctypes.c_int]
                user.GetClassNameW.argtypes=[wintypes.HWND,wintypes.LPWSTR,ctypes.c_int]
                user.EnumWindows.argtypes=[callback_type,wintypes.LPARAM]
                @callback_type
                def visit(hwnd,param):
                    title=ctypes.create_unicode_buffer(200);kind=ctypes.create_unicode_buffer(100)
                    user.GetWindowTextW(hwnd,title,200);user.GetClassNameW(hwnd,kind,100)
                    if title.value=='DBN Email Studio' and kind.value.startswith('Chrome_WidgetWin'):
                        result=ctypes.c_size_t()
                        for index,icon in enumerate(_icons):
                            if icon:send(hwnd,0x0080,index,icon,2,300,ctypes.byref(result))
                        try:_properties(hwnd,icon_path)
                        except (OSError,ValueError):pass
                    return True
                for attempt in range(12):
                    user.EnumWindows(visit,0);time.sleep(.5)
            except (OSError,ValueError,AttributeError):pass
            finally:
                try:ctypes.windll.ole32.CoUninitialize()
                except (OSError,AttributeError):pass
        _worker=threading.Thread(target=work,daemon=True);_worker.start()
