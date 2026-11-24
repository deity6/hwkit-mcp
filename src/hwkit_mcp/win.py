"""Windows 窗口枚举 + 非侵入式截图（纯 ctypes 标准库实现）。

设计约束（这些约束决定了下面每一处实现细节，注释里只写"为什么"）：

1. **非侵入**：绝不调用 SetForegroundWindow / SetFocus / SendInput / mouse_event /
   keybd_event / SetCursorPos。截图只走 ``PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT)``，
   由目标窗口自己把自己画到一个内存 DC 里，因此：
     - 不需要窗口前置、不需要可见于最上层；
     - 不会改变 Z 序、焦点、鼠标位置、键盘状态。
   只有在 PrintWindow 完全失败时才退回到 ``BitBlt`` 抓屏 + 窗口矩形裁剪（这条会受遮挡影响，
   结果里会标明用了哪条路径，不假装等价）。

2. **纯标准库**：不引入 pillow / pywin32。PNG 编码器自己写（zlib + struct，IHDR/IDAT/IEND
   三个 chunk 就够）。GDI+ 的 GdipSaveImageToFile 需要额外做 GdiplusStartup 的封送，
   而手写 PNG 只有 ~30 行且完全可控、无 GDI+ 的初始化顺序陷阱，所以选后者。

3. **DPI**：必须在任何窗口/屏幕 API 之前把进程设成 per-monitor DPI aware，否则在 125%
   缩放下 GetWindowRect / GetSystemMetrics 返回的是被系统缩放过的逻辑值，而 DIB 是物理
   像素，两者对不上就会出现"截出来的图被裁掉一圈"。这是进程级一次性设置，放在 import
   时用尽力而为的方式调用（失败不抛异常，只记在 ``DPI_AWARE`` 里供调用方判断）。

PNG 说明：DIB 是 32bpp BGRA，但 GDI 的 PrintWindow/BitBlt 不会写 alpha 通道（多半是 0），
所以 PNG 用 color type 2（truecolor，无 alpha）并把 alpha 直接丢掉——否则会得到一张
"全透明"的图，看图软件的棋盘格里什么也看不见。
"""

from __future__ import annotations

import ctypes
import struct
import time
import zlib
from ctypes import wintypes
from pathlib import Path

__all__ = [
    "list_windows",
    "find_window",
    "capture_window",
    "capture_screen",
    "png_size",
    "enable_dpi_awareness",
    "DPI_AWARE",
]

# --------------------------------------------------------------------------------------
# DLL 与常量
# --------------------------------------------------------------------------------------

user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
try:  # dwmapi 只用来过滤 UWP 的"幽灵窗口"，缺了也不影响主流程
    _dwmapi = ctypes.WinDLL("dwmapi", use_last_error=True)
except OSError:  # pragma: no cover - 只有极老/极精简的系统才会没有
    _dwmapi = None

PW_CLIENTONLY = 0x00000001
PW_RENDERFULLCONTENT = 0x00000002  # Win 8.1+：让窗口渲染完整内容（含 DWM 合成/DirectComposition）

SRCCOPY = 0x00CC0020
CAPTUREBLT = 0x40000000  # 连分层窗口一起抓；对 BitBlt 兜底路径有意义

BI_RGB = 0
DIB_RGB_COLORS = 0

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
DWMWA_CLOAKED = 14

# DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 —— 以 HANDLE(-4) 的形式传递
_DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)

_is_64bit = ctypes.sizeof(ctypes.c_void_p) == 8
_LONG_PTR = ctypes.c_ssize_t if _is_64bit else ctypes.c_long


# --------------------------------------------------------------------------------------
# 结构体
# --------------------------------------------------------------------------------------


class BITMAPINFOHEADER(ctypes.Structure):
    """GDI 的 BITMAPINFOHEADER；CreateDIBSection 需要它来描述目标位图格式。"""

    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    """BITMAPINFOHEADER + 调色板占位。32bpp 恒定色阶不需要调色板，但结构体大小要够。"""

    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", wintypes.DWORD * 3),
    ]


WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


# --------------------------------------------------------------------------------------
# 函数原型：必须显式声明，因为 HWND/HDC/HANDLE 是指针宽度，不声明 argtypes 时 ctypes 会按
# c_int 传参，64 位下高位被截断 —— 那是真实会崩的 bug，不是洁癖。
# --------------------------------------------------------------------------------------

user32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL

user32.IsWindow.argtypes = [wintypes.HWND]
user32.IsWindow.restype = wintypes.BOOL

user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.IsWindowVisible.restype = wintypes.BOOL

user32.IsIconic.argtypes = [wintypes.HWND]
user32.IsIconic.restype = wintypes.BOOL

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND

user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetWindowRect.restype = wintypes.BOOL

user32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetClientRect.restype = wintypes.BOOL

user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int

user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int

user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD

user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
user32.PrintWindow.restype = wintypes.BOOL

user32.GetDC.argtypes = [wintypes.HWND]
user32.GetDC.restype = wintypes.HDC

user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
user32.ReleaseDC.restype = ctypes.c_int

user32.GetSystemMetrics.argtypes = [ctypes.c_int]
user32.GetSystemMetrics.restype = ctypes.c_int

# 老 API，仅当新 API 不存在时才用；存在性在运行时判断，所以这里用 try/except 包住
try:
    user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
    user32.SetProcessDpiAwarenessContext.restype = wintypes.BOOL
except AttributeError:  # pragma: no cover - Win10 1703 之前
    pass

user32.SetProcessDPIAware.argtypes = []
user32.SetProcessDPIAware.restype = wintypes.BOOL

gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC

gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP

gdi32.CreateDIBSection.argtypes = [
    wintypes.HDC,
    ctypes.POINTER(BITMAPINFO),
    wintypes.UINT,
    ctypes.POINTER(ctypes.c_void_p),
    wintypes.HANDLE,
    wintypes.DWORD,
]
gdi32.CreateDIBSection.restype = wintypes.HBITMAP

gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ

gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteObject.restype = wintypes.BOOL

gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteDC.restype = wintypes.BOOL

gdi32.BitBlt.argtypes = [
    wintypes.HDC,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HDC,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.DWORD,
]
gdi32.BitBlt.restype = wintypes.BOOL

kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE

kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL

kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL

if _dwmapi is not None:
    _dwmapi.DwmGetWindowAttribute.argtypes = [
        wintypes.HWND,
        wintypes.DWORD,
        ctypes.c_void_p,
        wintypes.DWORD,
    ]
    _dwmapi.DwmGetWindowAttribute.restype = ctypes.c_long  # HRESULT


# --------------------------------------------------------------------------------------
# DPI
# --------------------------------------------------------------------------------------


def enable_dpi_awareness() -> bool:
    """把当前进程设成 per-monitor DPI aware（V2 优先，逐级回退）。

    为什么必须先做：在 125% 缩放下，非 DPI-aware 进程看到的是一张被系统缩放过的"逻辑
    坐标系"，GetWindowRect 给出的是逻辑像素，而我们分配的 DIB 是物理像素。两者混用会让
    截图内容与实际尺寸错位（典型症状：右边和下边被裁掉、或者右下角出现黑边）。

    返回:
        True 表示至少有一级设置成功；False 表示系统拒绝了全部尝试（通常意味着
        进程清单里已经写死了 DPI 声明，此时必须以那个声明为准，尺寸仍可能一致）。

    注意：这是进程级、不可逆的操作，重复调用无副作用。
    """
    # 1) Win10 1703+ 的 PER_MONITOR_AWARE_V2：能正确处理跨屏拖动时的 DPI 变化
    try:
        if user32.SetProcessDpiAwarenessContext(_DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2):
            return True
    except (AttributeError, OSError):
        pass

    # 2) Win8.1+ 的 shcore.SetProcessDpiAwareness(2) = PROCESS_PER_MONITOR_DPI_AWARE
    try:
        shcore = ctypes.WinDLL("shcore", use_last_error=True)
        shcore.SetProcessDpiAwareness.argtypes = [ctypes.c_int]
        shcore.SetProcessDpiAwareness.restype = ctypes.c_long
        # 只调一次：这个 API 第一次调用就决定一切，重复调用只会返回同样的拒绝原因
        hr = shcore.SetProcessDpiAwareness(2) & 0xFFFFFFFF
        # S_OK=0；E_ACCESSDENIED=0x80070005 表示清单里已经声明过 aware，等价于"已达成目的"
        if hr in (0, 0x80070005):
            return True
    except (AttributeError, OSError):
        pass

    # 3) 老 API：system DPI aware，聊胜于无
    try:
        if user32.SetProcessDPIAware():
            return True
    except (AttributeError, OSError):
        pass
    return False


DPI_AWARE: bool = enable_dpi_awareness()


# --------------------------------------------------------------------------------------
# 进程信息
# --------------------------------------------------------------------------------------


def _process_name(pid: int) -> str:
    """用 QueryFullProcessImageNameW 取进程可执行文件名（不含目录）。

    为什么不用 GetModuleFileNameEx：那个需要 PROCESS_VM_READ 权限，对别的用户/受保护进程
    会失败；QueryFullProcessImageName 只要 PROCESS_QUERY_LIMITED_INFORMATION，覆盖面大得多。
    失败时返回空串而不是抛异常——枚举窗口不该因为一个取不到的进程名而中断。
    """
    if not pid:
        return ""
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buf))
        if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return Path(buf.value).name
        return ""
    finally:
        kernel32.CloseHandle(handle)


def _window_title(hwnd: int) -> str:
    """取窗口标题。先问长度再分配，避免固定大缓冲区的截断问题。"""
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    copied = user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value[:copied] if copied else ""


def _window_class(hwnd: int) -> str:
    """取窗口类名（如 Chrome_WidgetWin_1 / CabinetWClass）。"""
    buf = ctypes.create_unicode_buffer(512)
    copied = user32.GetClassNameW(hwnd, buf, 512)
    return buf.value[:copied] if copied else ""


def _window_rect(hwnd: int) -> dict:
    """取窗口矩形（屏幕坐标、物理像素）。失败返回全 0 的矩形，让调用方自己判零。"""
    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return {"x": 0, "y": 0, "width": 0, "height": 0}
    return {
        "x": int(rect.left),
        "y": int(rect.top),
        "width": int(rect.right - rect.left),
        "height": int(rect.bottom - rect.top),
    }


def _is_cloaked(hwnd: int) -> bool:
    """判断窗口是否被 DWM "cloaked"。

    为什么需要：Win10/11 上每个 UWP/Store 应用都会留下一个尺寸巨大、标题正常的顶层窗口，
    但它被 DWM 藏起来了（cloaked=1）。EnumWindows 会把它列出来，导致枚举结果里混进一堆
    根本看不见、也截不出内容的幽灵窗口。
    """
    if _dwmapi is None:
        return False
    value = wintypes.DWORD(0)
    hr = _dwmapi.DwmGetWindowAttribute(
        hwnd, DWMWA_CLOAKED, ctypes.byref(value), ctypes.sizeof(value)
    )
    return hr == 0 and value.value != 0


# --------------------------------------------------------------------------------------
# 枚举
# --------------------------------------------------------------------------------------


def _enum_top_windows(visible_only: bool, skip_cloaked: bool) -> list[dict]:
    """底层枚举：一次性走完 EnumWindows，返回所有顶层窗口的描述。"""
    results: list[dict] = []
    foreground = int(user32.GetForegroundWindow() or 0)
    name_cache: dict[int, str] = {}

    def _callback(hwnd, _lparam):
        try:
            # 可见性过滤放在最前面：不可见窗口占绝大多数，先剔掉省后面的 API 调用
            if visible_only and not user32.IsWindowVisible(hwnd):
                return True
            if skip_cloaked and _is_cloaked(hwnd):
                return True

            pid_value = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_value))
            pid = int(pid_value.value)

            # 同一个进程往往有多个窗口，缓存一次进程名，避免反复 OpenProcess
            if pid in name_cache:
                proc = name_cache[pid]
            else:
                proc = _process_name(pid)
                name_cache[pid] = proc

            results.append(
                {
                    "hwnd": int(hwnd),
                    "pid": pid,
                    "title": _window_title(hwnd),
                    "class": _window_class(hwnd),
                    "rect": _window_rect(hwnd),
                    "minimized": bool(user32.IsIconic(hwnd)),
                    "foreground": int(hwnd) == foreground,
                    "process": proc,
                }
            )
        except Exception:  # noqa: BLE001 - 回调里抛异常会让整个枚举静默中断，宁可跳过这一项
            pass
        return True

    user32.EnumWindows(WNDENUMPROC(_callback), 0)
    return results


def list_windows(visible_only: bool = True) -> list[dict]:
    """枚举顶层窗口。

    参数:
        visible_only: True 时只返回 IsWindowVisible 为真的窗口，并额外剔除 DWM cloaked
            的幽灵窗口（UWP 壳）。False 时返回全部顶层窗口，包含隐藏窗口和 cloaked 窗口。

    返回:
        列表，每项为::

            {"hwnd": int, "pid": int, "title": str, "class": str,
             "rect": {"x","y","width","height"},
             "minimized": bool, "foreground": bool, "process": str}

        其中 ``rect`` 是 GetWindowRect 的物理像素（含不可见边框），``process`` 是进程可执行
        文件名（不含目录），``foreground`` 表示是否是当前前台窗口。

    说明: 本函数只做只读查询，不激活、不排序、不发送任何输入事件。
    """
    return _enum_top_windows(visible_only=visible_only, skip_cloaked=visible_only)


def find_window(title_contains: str = "", pid: int = 0) -> dict | None:
    """按标题子串或 pid 找第一个匹配窗口。

    参数:
        title_contains: 标题子串，大小写不敏感；空串表示不按标题过滤。
        pid: 进程 ID；0 表示不按 pid 过滤。两者都给时要求同时满足。

    返回:
        与 :func:`list_windows` 同结构的一项；找不到返回 None。

    匹配策略（这是"找第一个"该有的优先级，不是随手排序）：
        1. 先在可见窗口里按可见性原序找，优先非最小化、标题非空的窗口；
        2. 可见窗口里找不到，再退到隐藏窗口全表查一遍 —— 调用方给 pid 时经常指的是一个
           刚好最小化/隐藏的窗口。
    """
    needle = title_contains.lower()

    def _matches(win: dict) -> bool:
        if pid and win["pid"] != pid:
            return False
        if needle and needle not in win["title"].lower():
            return False
        return True

    def _pick(windows: list[dict]) -> dict | None:
        matched = [w for w in windows if _matches(w)]
        if not matched:
            return None
        # 排序但不改原枚举顺序的语义：优先"能截的"窗口（非最小化、标题非空、面积大）
        matched.sort(
            key=lambda w: (
                not w["minimized"],
                bool(w["title"]),
                w["rect"]["width"] * w["rect"]["height"],
            ),
            reverse=True,
        )
        return matched[0]

    found = _pick(_enum_top_windows(visible_only=True, skip_cloaked=True))
    if found is not None:
        return found
    return _pick(_enum_top_windows(visible_only=False, skip_cloaked=False))


# --------------------------------------------------------------------------------------
# PNG 编码（手写，只依赖 zlib/struct）
# --------------------------------------------------------------------------------------

_PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _png_chunk(tag: bytes, data: bytes) -> bytes:
    """拼一个 PNG chunk：长度 + 类型 + 数据 + CRC32(类型+数据)。"""
    return (
        struct.pack(">I", len(data))
        + tag
        + data
        + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    )


def _encode_png(rgb_rows: bytes, width: int, height: int, level: int = 6) -> bytes:
    """把"每行 width*3 字节依次相连"的 RGB 数据编码成 PNG（8bit truecolor，无 alpha）。

    参数:
        rgb_rows: 长度必须是 width*height*3，行序自上而下。
        width / height: 像素尺寸。
        level: zlib 压缩等级；截图场景 6 已经足够（9 只多几个百分点，慢一倍）。

    为什么用 color type 2 而不是 6：GDI 写 DIB 时不碰 alpha 通道，PrintWindow/BitBlt 之后
    alpha 基本恒为 0。写成带 alpha 的 PNG 在多数看图器/浏览器里会渲染成全透明，看起来像
    "抓空了"。直接丢掉 alpha 得到的就是所见即所得的画面。

    为什么只用一个 IDAT：单帧截图的数据量（2560x1440 约 11MB 原始）远小于任何实现会遇到的
    内存限制，拆多个 IDAT 只是给流式编码用的，这里不需要。
    """
    expected = width * height * 3
    if len(rgb_rows) != expected:
        raise ValueError(f"RGB 数据长度不符: {len(rgb_rows)} != {expected}")

    ihdr = struct.pack(
        ">IIBBBBB",
        width,
        height,
        8,  # bit depth
        2,  # color type: truecolor (RGB)
        0,  # compression: deflate
        0,  # filter: adaptive
        0,  # interlace: none
    )

    # 每行前面塞一个 filter byte=0（None）。截图是逐行独立压缩的，用 Paeth/Sub 之类的
    # 自适应过滤收益不大，反而让解码端（包括测试脚本）复杂化。
    stride = width * 3
    raw = bytearray(expected + height)
    pos = 0
    for y in range(height):
        src = y * stride
        raw[pos] = 0
        raw[pos + 1 : pos + 1 + stride] = rgb_rows[src : src + stride]
        pos += stride + 1

    return (
        _PNG_SIG
        + _png_chunk(b"IHDR", ihdr)
        + _png_chunk(b"IDAT", zlib.compress(bytes(raw), level))
        + _png_chunk(b"IEND", b"")
    )


def _dib_bits_to_rgb(bits_ptr: int, width: int, height: int) -> bytes:
    """把 DIB section 的 BGRA 缓冲转成自上而下的紧凑 RGB 字节串。

    两个细节：
      * DIB 是**自下而上**存的（biHeight 为正时），所以行序要翻转；
      * 32bpp 的行跨度恰好是 width*4，不需要处理对齐填充。
    alpha 用扩展切片 ``del buf[3::4]`` 一次剔掉——这是 C 速度的，比 Python 循环快两个数量级
    （2560x1440 用逐像素循环要好几秒）。
    """
    stride = width * 4
    buf = ctypes.string_at(bits_ptr, stride * height)

    if height > 1:
        rows = [buf[i : i + stride] for i in range(0, stride * height, stride)]
        rows.reverse()
        buf = b"".join(rows)

    arr = bytearray(buf)
    del arr[3::4]  # BGRA -> BGR
    # BGR -> RGB：交换每 3 字节里的第 0、2 位。同样用切片做，避免逐像素（快 ~50x）
    arr[0::3], arr[2::3] = arr[2::3], arr[0::3]
    return bytes(arr)


# --------------------------------------------------------------------------------------
# DIB 辅助
# --------------------------------------------------------------------------------------


class _MemDib:
    """一块 32bpp 的自上而下可读 DIB section 与其内存 DC。

    为什么用 CreateDIBSection 而不是 CreateCompatibleBitmap：后者得到的是设备相关位图，
    GetDIBits 之外没法直接读像素；DIB section 直接给一个进程内可读的缓冲区指针，
    PrintWindow / BitBlt 画进去之后立刻就能 ``string_at`` 取到，少一次整幅拷贝。
    """

    def __init__(self, width: int, height: int) -> None:
        self.width = width
        self.height = height
        self.hdc = gdi32.CreateCompatibleDC(None)
        if not self.hdc:
            raise OSError("CreateCompatibleDC 失败")

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = height  # 正值 = 自下而上，与 GDI 惯例一致
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB

        self.bits = ctypes.c_void_p()
        self.hbitmap = gdi32.CreateDIBSection(
            self.hdc,
            ctypes.byref(bmi),
            DIB_RGB_COLORS,
            ctypes.byref(self.bits),
            None,
            0,
        )
        if not self.hbitmap or not self.bits:
            gdi32.DeleteDC(self.hdc)
            raise OSError("CreateDIBSection 失败")

        self.old = gdi32.SelectObject(self.hdc, self.hbitmap)

    def rgb(self) -> bytes:
        """读出像素并转成自上而下 RGB。"""
        return _dib_bits_to_rgb(self.bits.value, self.width, self.height)

    def close(self) -> None:
        """按创建的反序释放，避免 GDI 对象泄漏（长时间跑 MCP 服务会累积）。"""
        try:
            if self.old:
                gdi32.SelectObject(self.hdc, self.old)
        finally:
            if self.hbitmap:
                gdi32.DeleteObject(self.hbitmap)
                self.hbitmap = None
            if self.hdc:
                gdi32.DeleteDC(self.hdc)
                self.hdc = None

    def __enter__(self) -> "_MemDib":
        return self

    def __exit__(self, *_exc) -> None:
        self.close()


def _default_shot_path(hwnd: int) -> Path:
    """默认输出路径：%TEMP%\\hwkit_shot_<hwnd>_<时间戳>.png。

    为什么落到 TEMP：调用方没指定路径时这是唯一"一定可写"的位置，且不会把临时截图
    丢进用户的工作目录里。
    """
    import os
    import tempfile

    temp = os.environ.get("TEMP") or tempfile.gettempdir()
    stamp = time.strftime("%Y%m%d_%H%M%S") + f"_{int(time.time() * 1000) % 1000:03d}"
    return Path(temp) / f"hwkit_shot_{hwnd}_{stamp}.png"


def _write_png(path: Path, rgb: bytes, width: int, height: int) -> int:
    """编码并落盘，返回字节数。父目录不存在时自动建。"""
    data = _encode_png(rgb, width, height)
    path.parent.mkdir(parents=True, exist_ok=True)
    # 二进制写：PNG 不是文本，任何 encoding 参数都会毁掉它
    path.write_bytes(data)
    return len(data)


def _sample_distinct_colors(rgb: bytes, limit: int = 250_000) -> int:
    """抽样统计不同 RGB 值的个数，用来做"这张图是不是纯色/全黑"的快速自检。

    为什么抽样：逐像素 set 在全屏截图（370 万像素）上要 2~4 秒纯 Python 开销，
    而判断"有没有内容"根本不需要全量。抽样步长使样本数不超过 limit。
    """
    pixels = len(rgb) // 3
    if pixels == 0:
        return 0
    step = max(1, pixels // limit)
    seen = set()
    for i in range(0, pixels, step):
        off = i * 3
        seen.add(rgb[off : off + 3])
    return len(seen)


# --------------------------------------------------------------------------------------
# 截图
# --------------------------------------------------------------------------------------


def capture_window(
    title_contains: str = "",
    pid: int = 0,
    out_path: str = "",
    hwnd: int = 0,
) -> dict:
    """非侵入式抓取窗口位图并存成 PNG。

    参数:
        title_contains: 标题子串（配合 find_window 使用；hwnd 非 0 时忽略）。
        pid: 目标进程 ID（同上）。
        out_path: 输出 PNG 路径；空串时写到 ``%TEMP%\\hwkit_shot_<hwnd>_<时间戳>.png``。
        hwnd: 直接指定窗口句柄，优先级最高。

    返回:
        ``{"ok": bool, "path": str, "width": int, "height": int, "bytes": int, "hwnd": int}``
        失败时 ``ok=False``、``width/height/bytes`` 为 0，并额外带 ``error`` 说明原因；
        成功时额外带 ``method``（用了哪条抓取路径）和 ``distinct_colors``（抽样的颜色数）。

    实现要点:
        * 尺寸取 GetWindowRect（物理像素，含不可见边框）—— 与 list_windows 的 rect 同源，
          这样调用方拿 rect 和 PNG 尺寸能直接对上；
        * 最小化窗口直接失败：最小化后窗口没有有效的绘制表面，PrintWindow 只会返回一张
          空白/残留位图，假成功比失败更有害；
        * 首选 ``PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT=2)``：由目标窗口自己绘制，
          窗口被遮挡或不在最前也能拿到完整内容，且完全不动窗口状态；
        * 2 失败退到 ``PrintWindow(hwnd, hdc, 0)``（老式整窗打印，某些自绘程序只支持这个）；
        * 再失败才退到 BitBlt 抓屏裁剪 —— 这条**受遮挡影响**，结果里会如实标 method。
        * 全程不调用 SetForegroundWindow/SetFocus/SendInput，不移动鼠标。
    """
    hwnd = int(hwnd or 0)
    resolved_by = "hwnd"
    if not hwnd:
        found = find_window(title_contains=title_contains, pid=pid)
        if found is None:
            return {
                "ok": False,
                "path": "",
                "width": 0,
                "height": 0,
                "bytes": 0,
                "hwnd": 0,
                "error": f"找不到匹配窗口 (title_contains={title_contains!r}, pid={pid})",
            }
        hwnd = found["hwnd"]
        resolved_by = "find_window"

    if not user32.IsWindow(hwnd):
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "hwnd": hwnd,
            "error": f"hwnd {hwnd} 不是有效窗口",
        }

    if user32.IsIconic(hwnd):
        # 最小化窗口没有可绘制表面；这里明确报失败，而不是返回一张黑图骗调用方
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "hwnd": hwnd,
            "error": "窗口已最小化(IsIconic)，没有可绘制表面，无法非侵入式截图",
        }

    rect = _window_rect(hwnd)
    width, height = rect["width"], rect["height"]
    if width <= 0 or height <= 0:
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "hwnd": hwnd,
            "error": f"窗口矩形无效: {rect}",
        }

    method = ""
    try:
        with _MemDib(width, height) as dib:
            # 第 1 选择：PW_RENDERFULLCONTENT，画的是窗口的完整合成内容
            if user32.PrintWindow(hwnd, dib.hdc, PW_RENDERFULLCONTENT):
                method = "PrintWindow:PW_RENDERFULLCONTENT"
            # 第 2 选择：老式 PrintWindow(0)
            elif user32.PrintWindow(hwnd, dib.hdc, 0):
                method = "PrintWindow:legacy"
            else:
                # 第 3 选择：从屏幕 DC 直接复制过来再裁剪。会受遮挡影响，且要求窗口可见。
                screen_dc = user32.GetDC(None)
                if not screen_dc:
                    return {
                        "ok": False,
                        "path": "",
                        "width": 0,
                        "height": 0,
                        "bytes": 0,
                        "hwnd": hwnd,
                        "error": "PrintWindow 失败，且 GetDC(NULL) 也失败",
                    }
                try:
                    ok = gdi32.BitBlt(
                        dib.hdc,
                        0,
                        0,
                        width,
                        height,
                        screen_dc,
                        rect["x"],
                        rect["y"],
                        SRCCOPY | CAPTUREBLT,
                    )
                finally:
                    user32.ReleaseDC(None, screen_dc)
                if not ok:
                    return {
                        "ok": False,
                        "path": "",
                        "width": 0,
                        "height": 0,
                        "bytes": 0,
                        "hwnd": hwnd,
                        "error": "PrintWindow(2/0) 与 BitBlt 兜底全部失败",
                    }
                method = "BitBlt:screen-crop"

            rgb = dib.rgb()
    except OSError as exc:
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "hwnd": hwnd,
            "error": f"GDI 初始化失败: {exc}",
        }

    target = Path(out_path) if out_path else _default_shot_path(hwnd)
    nbytes = _write_png(target, rgb, width, height)
    colors = _sample_distinct_colors(rgb)

    return {
        "ok": True,
        "path": str(target),
        "width": width,
        "height": height,
        "bytes": nbytes,
        "hwnd": hwnd,
        "method": method,
        "distinct_colors": colors,
        # 启发式告警：某些程序（极少见，通常是自绘 + 硬件合成表面）会让 PrintWindow 返回一张
        # 纯色位图，调用方 API 层看起来"成功了"但内容其实是空的。给出信号而不是替调用方决定。
        "blank_like": colors <= 2,
        "resolved_by": resolved_by,
        "dpi_aware": DPI_AWARE,
    }


def capture_screen(region: dict | None = None, out_path: str = "") -> dict:
    """抓整个屏幕或指定区域并存成 PNG。

    参数:
        region: ``{"x","y","width","height"}``；None 表示抓整个主屏。
        out_path: 输出路径；空串时写到 ``%TEMP%\\hwkit_shot_screen_<时间戳>.png``。

    返回:
        ``{"ok","path","width","height","bytes","x","y","screen_width","screen_height"}``，
        失败时 ``ok=False`` 并带 ``error``。

    实现要点:
        * 用 ``GetDC(NULL)`` 取屏幕 DC，再 ``BitBlt`` 到一块自建的 32bpp DIB —— 全程只读，
          不改动任何窗口状态（不截图就先 RestoreWindow 那类"化妆"操作一概没有）；
        * ``CAPTUREBLT`` 让分层窗口（部分带透明/合成的窗口）也一起进画面；
        * 区域越界时裁到屏幕范围内，而不是返回一张带黑边的图；
        * 尺寸取 GetSystemMetrics(0/1)：因为进程已 per-monitor DPI aware，这里拿到的是
          主屏物理像素，与 DIB 尺寸口径一致。
    """
    screen_w = int(user32.GetSystemMetrics(0))
    screen_h = int(user32.GetSystemMetrics(1))
    if screen_w <= 0 or screen_h <= 0:
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "error": f"GetSystemMetrics 返回异常尺寸 {screen_w}x{screen_h}",
        }

    if region:
        x = int(region.get("x", 0))
        y = int(region.get("y", 0))
        width = int(region.get("width", screen_w - x))
        height = int(region.get("height", screen_h - y))
        # 裁剪到屏幕内：越界区域在 BitBlt 里会得到未定义内容，不如老实裁掉
        x = max(0, min(x, screen_w - 1))
        y = max(0, min(y, screen_h - 1))
        width = max(1, min(width, screen_w - x))
        height = max(1, min(height, screen_h - y))
    else:
        x = y = 0
        width, height = screen_w, screen_h

    screen_dc = user32.GetDC(None)
    if not screen_dc:
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "error": "GetDC(NULL) 失败",
        }
    try:
        with _MemDib(width, height) as dib:
            if not gdi32.BitBlt(
                dib.hdc, 0, 0, width, height, screen_dc, x, y, SRCCOPY | CAPTUREBLT
            ):
                return {
                    "ok": False,
                    "path": "",
                    "width": 0,
                    "height": 0,
                    "bytes": 0,
                    "error": "BitBlt 从屏幕 DC 复制失败",
                }
            rgb = dib.rgb()
    except OSError as exc:
        return {
            "ok": False,
            "path": "",
            "width": 0,
            "height": 0,
            "bytes": 0,
            "error": f"GDI 初始化失败: {exc}",
        }
    finally:
        user32.ReleaseDC(None, screen_dc)

    if out_path:
        target = Path(out_path)
    else:
        import os
        import tempfile

        temp = os.environ.get("TEMP") or tempfile.gettempdir()
        stamp = time.strftime("%Y%m%d_%H%M%S") + f"_{int(time.time() * 1000) % 1000:03d}"
        target = Path(temp) / f"hwkit_shot_screen_{stamp}.png"

    nbytes = _write_png(target, rgb, width, height)
    return {
        "ok": True,
        "path": str(target),
        "width": width,
        "height": height,
        "bytes": nbytes,
        "x": x,
        "y": y,
        "screen_width": screen_w,
        "screen_height": screen_h,
        "distinct_colors": _sample_distinct_colors(rgb),
        "dpi_aware": DPI_AWARE,
    }


# --------------------------------------------------------------------------------------
# PNG 头解析（供测试/调用方校验用）
# --------------------------------------------------------------------------------------


def png_size(path: str) -> dict:
    """读 PNG 的 IHDR 头返回 ``{"width","height"}``。

    只读前 33 字节（8 字节签名 + 4 长度 + 4 类型 + 13 数据 + 4 CRC），不解码像素——
    这是一个廉价的"文件是不是真 PNG、尺寸对不对"的校验入口。

    参数:
        path: PNG 文件路径（支持含中文/全角字符的路径；严格按二进制读，不做任何编码转换）。

    异常:
        ValueError: 文件短于 33 字节、或签名/IHDR 不匹配。
    """
    p = Path(path)
    with p.open("rb") as fh:
        head = fh.read(33)
    if len(head) < 33:
        raise ValueError(f"{p} 太短，不是有效 PNG（{len(head)} 字节）")
    if head[:8] != _PNG_SIG:
        raise ValueError(f"{p} 不是 PNG（签名不符）")
    if head[12:16] != b"IHDR":
        raise ValueError(f"{p} 首个 chunk 不是 IHDR")
    width, height = struct.unpack(">II", head[16:24])
    return {"width": int(width), "height": int(height)}
