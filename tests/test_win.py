"""hwkit_mcp.win 的自测脚本（纯标准库，含一个最小 PNG 解码器用于验证"图里真有内容"）。

运行:
    python tests\test_win.py

覆盖:
    1. list_windows() 非空，且存在标题非空的窗口（打印前 5 个）
    2. capture_window() 抓真实窗口成功；PNG 尺寸≈窗口 rect(±2px)；解 IDAT 统计不同颜色数 > 50
    3. capture_screen() 成功；尺寸 == getSystemMetrics(0/1)
    4. 截图前后前台窗口不变、鼠标位置不变（证明非侵入）

所有文件读写显式 encoding / 二进制，路径用 pathlib.Path（项目路径含中文和全角括号）。
"""

from __future__ import annotations

import ctypes
import struct
import sys
import zlib
from ctypes import wintypes
from pathlib import Path

TEST_DIR = Path(__file__).resolve().parent
PROJECT = TEST_DIR.parent
SRC = PROJECT / "src"
sys.path.insert(0, str(SRC))

from hwkit_mcp import win  # noqa: E402

# 测试产物统一放项目下的 _shots 目录，便于人工复核；不污染 TEMP 里的默认路径测试
OUT_DIR = PROJECT / "_shots"

_passed: list[str] = []
_failed: list[str] = []


def check(name: str, cond: bool, detail: str = "") -> bool:
    """记录一条断言结果并打印。不中断执行——一次跑完拿到全部结论比"第一条失败就退出"有用。"""
    if cond:
        _passed.append(name)
        print(f"  [PASS] {name}" + (f" — {detail}" if detail else ""))
    else:
        _failed.append(name)
        print(f"  [FAIL] {name}" + (f" — {detail}" if detail else ""))
    return cond


# --------------------------------------------------------------------------------------
# 最小 PNG 解码器：只为验证"像素里真有内容"，不追求完整性
# --------------------------------------------------------------------------------------


def decode_png(path: Path) -> tuple[int, int, int, bytes]:
    """解出 (width, height, bytes_per_pixel, 自上而下像素字节)。

    支持 8bit、color type 2(RGB)/6(RGBA)、全部 5 种 scanline filter。
    必须实现反过滤而不是只剥 filter byte：即使本模块只用 filter=0，解码器也要能独立
    验证一个任意 PNG —— 否则"解码器只认得自己编码器的偷懒输出"会让测试失去意义。
    """
    data = path.read_bytes()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("PNG 签名不符")

    pos = 8
    idat = bytearray()
    width = height = depth = ctype = None
    while pos < len(data):
        (length,) = struct.unpack(">I", data[pos : pos + 4])
        tag = data[pos + 4 : pos + 8]
        body = data[pos + 8 : pos + 8 + length]
        pos += 12 + length  # 4 len + 4 tag + body + 4 crc
        if tag == b"IHDR":
            width, height, depth, ctype = struct.unpack(">IIBB", body[:10])
        elif tag == b"IDAT":
            idat += body
        elif tag == b"IEND":
            break

    if depth != 8 or ctype not in (2, 6):
        raise ValueError(f"仅支持 8bit RGB/RGBA，实际 depth={depth} colortype={ctype}")

    bpp = 3 if ctype == 2 else 4
    stride = width * bpp
    raw = zlib.decompress(bytes(idat))
    out = bytearray(stride * height)
    prev = bytearray(stride)
    p = 0
    for y in range(height):
        ftype = raw[p]
        p += 1
        line = bytearray(raw[p : p + stride])
        p += stride
        if ftype == 1:  # Sub
            for i in range(bpp, stride):
                line[i] = (line[i] + line[i - bpp]) & 0xFF
        elif ftype == 2:  # Up
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 0xFF
        elif ftype == 3:  # Average
            for i in range(stride):
                left = line[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif ftype == 4:  # Paeth
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                b = prev[i]
                c = prev[i - bpp] if i >= bpp else 0
                pa, pb, pc = abs(b - c), abs(a - c), abs(a + b - 2 * c)
                pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[i] = (line[i] + pr) & 0xFF
        elif ftype != 0:
            raise ValueError(f"未知 filter type {ftype}")
        out[y * stride : (y + 1) * stride] = line
        prev = line
    return width, height, bpp, bytes(out)


def distinct_colors(pixels: bytes, bpp: int, max_samples: int = 300_000) -> int:
    """抽样统计不同 RGB 值个数（步长自适应，控制纯 Python 开销）。"""
    count = len(pixels) // bpp
    if count == 0:
        return 0
    step = max(1, count // max_samples)
    seen: set[bytes] = set()
    for i in range(0, count, step):
        off = i * bpp
        seen.add(pixels[off : off + 3])
    return len(seen)


# --------------------------------------------------------------------------------------
# 非侵入性证据采集（直接调 user32，不依赖被测模块的实现细节）
# --------------------------------------------------------------------------------------

_u32 = ctypes.WinDLL("user32", use_last_error=True)
_u32.GetForegroundWindow.restype = wintypes.HWND
_u32.GetForegroundWindow.argtypes = []


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


_u32.GetCursorPos.argtypes = [ctypes.POINTER(_POINT)]
_u32.GetCursorPos.restype = wintypes.BOOL


def cursor_pos() -> tuple[int, int]:
    """取当前鼠标位置（用于证明截图过程没有移动鼠标）。"""
    pt = _POINT()
    _u32.GetCursorPos(ctypes.byref(pt))
    return int(pt.x), int(pt.y)


def foreground() -> int:
    """取当前前台窗口句柄。"""
    return int(_u32.GetForegroundWindow() or 0)


# --------------------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------------------


def main() -> int:
    print("=" * 78)
    print("hwkit_mcp.win 自测")
    print(f"  python      : {sys.version.split()[0]}  ({sys.executable})")
    print(f"  模块        : {win.__file__}")
    print(f"  DPI_AWARE   : {win.DPI_AWARE}")
    print(f"  systemMetrics: {win.user32.GetSystemMetrics(0)} x {win.user32.GetSystemMetrics(1)}")
    print("=" * 78)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # ---------------- 测试 1: list_windows ----------------
    print("\n[1] list_windows()")
    windows = win.list_windows()
    check("list_windows 返回非空", len(windows) > 0, f"共 {len(windows)} 个可见顶层窗口")

    titled = [w for w in windows if w["title"].strip()]
    check("存在标题非空的窗口", len(titled) > 0, f"{len(titled)} 个有标题")

    print("  前 5 个窗口:")
    for w in windows[:5]:
        r = w["rect"]
        print(
            f"    hwnd={w['hwnd']:<10} pid={w['pid']:<6} "
            f"{r['width']}x{r['height']:<5} minimized={w['minimized']!s:<5} "
            f"fg={w['foreground']!s:<5} proc={w['process']:<22} "
            f"class={w['class'][:28]:<28} title={w['title'][:44]!r}"
        )

    # 结构检查：字段齐全且类型正确
    required = {"hwnd", "pid", "title", "class", "rect", "minimized", "foreground", "process"}
    check("每项字段齐全", all(required <= set(w) for w in windows))
    check(
        "rect 四键齐全且为 int",
        all(
            set(w["rect"]) == {"x", "y", "width", "height"}
            and all(isinstance(v, int) for v in w["rect"].values())
            for w in windows
        ),
    )

    # ---------------- 测试 2: capture_window ----------------
    print("\n[2] capture_window()")
    # 候选：有标题、没最小化、面积够大。
    # 为什么排除这几个类名：Progman/WorkerW 是桌面壁纸窗口，Shell_TrayWnd 是任务栏，
    # CoreWindow 是 UWP 壳。它们虽然也能截到东西，但用它们"证明抓到了真实应用内容"
    # 说服力不足——真实场景要截的是被挡住的应用窗口。
    _SHELL_CLASSES = {"Progman", "WorkerW", "Shell_TrayWnd", "Windows.UI.Core.CoreWindow",
                      "XamlExplorerHostIslandWindow"}
    candidates = [
        w
        for w in titled
        if not w["minimized"]
        and w["rect"]["width"] >= 300
        and w["rect"]["height"] >= 200
        and w["class"] not in _SHELL_CLASSES
    ]
    # 优先非前台窗口：这样"截到了完整内容"就同时证明了不依赖窗口前置（可抓被遮挡的窗口）
    candidates.sort(
        key=lambda w: (
            not w["foreground"],
            w["rect"]["width"] * w["rect"]["height"],
        ),
        reverse=True,
    )
    print(f"  候选窗口 {len(candidates)} 个（已排除桌面/任务栏壳窗口），按面积降序尝试前 6 个")

    fg_before_hwnd = foreground()
    cursor_before = cursor_pos()
    fg_flags_before = {w["hwnd"]: w["foreground"] for w in windows}

    shot: dict | None = None
    target_win: dict | None = None
    tried: list[str] = []
    for cand in candidates[:6]:
        out = OUT_DIR / f"shot_hwnd{cand['hwnd']}.png"
        res = win.capture_window(hwnd=cand["hwnd"], out_path=str(out))
        tried.append(f"hwnd={cand['hwnd']} ok={res.get('ok')} method={res.get('method')}")
        if not res.get("ok"):
            print(f"    跳过 hwnd={cand['hwnd']}: {res.get('error')}")
            continue
        # 现场验内容：纯色/全黑的图不能算成功，必须换下一个候选
        w_png, h_png, bpp, pixels = decode_png(Path(res["path"]))
        colors = distinct_colors(pixels, bpp)
        print(
            f"    尝试 hwnd={cand['hwnd']} {cand['rect']['width']}x{cand['rect']['height']} "
            f"-> PNG {w_png}x{h_png} colors={colors} method={res.get('method')} "
            f"title={cand['title'][:40]!r}"
        )
        if colors > 50:
            shot = {**res, "_png": (w_png, h_png, bpp, pixels), "_colors": colors}
            target_win = cand
            break

    check("capture_window 返回 ok=True", shot is not None, "; ".join(tried) or "无候选")

    if shot and target_win:
        path = Path(shot["path"])
        check("PNG 文件存在", path.exists(), str(path))
        check("返回的 bytes 与实际文件大小一致", path.stat().st_size == shot["bytes"],
              f"{shot['bytes']} bytes")

        reported = win.png_size(str(path))
        print(f"    png_size() = {reported}")
        check("png_size 与解码结果一致",
              reported["width"] == shot["_png"][0] and reported["height"] == shot["_png"][1])

        rect = target_win["rect"]
        dw = abs(shot["width"] - rect["width"])
        dh = abs(shot["height"] - rect["height"])
        check(
            "PNG 宽高与窗口 rect 一致(±2px)",
            dw <= 2 and dh <= 2 and reported["width"] == shot["width"] and reported["height"] == shot["height"],
            f"png={reported['width']}x{reported['height']} rect={rect['width']}x{rect['height']} "
            f"diff=({dw},{dh})",
        )
        check("图片不是纯色空白(不同颜色 > 50)", shot["_colors"] > 50, f"{shot['_colors']} 种颜色")
        print(f"    抓到窗口: hwnd={shot['hwnd']} title={target_win['title']!r} "
              f"class={target_win['class']} proc={target_win['process']}")

    # capture_window 的两种寻址方式 + 最小化/找不到 的错误分支
    print("\n[2b] capture_window 其它寻址/错误分支")
    if target_win:
        by_pid = win.capture_window(pid=target_win["pid"], out_path=str(OUT_DIR / "shot_by_pid.png"))
        check("按 pid 抓取成功", by_pid.get("ok") is True,
              f"hwnd={by_pid.get('hwnd')} via={by_pid.get('resolved_by')}")
        keyword = target_win["title"][:8]
        by_title = win.capture_window(title_contains=keyword, pid=target_win["pid"],
                                      out_path=str(OUT_DIR / "shot_by_title.png"))
        check(f"按标题子串 {keyword!r} 抓取成功", by_title.get("ok") is True,
              f"hwnd={by_title.get('hwnd')}")

    # out_path 留空时必须落到 %TEMP%\hwkit_shot_<hwnd>_<时间戳>.png
    if target_win:
        import os
        import tempfile

        dflt = win.capture_window(hwnd=target_win["hwnd"])
        temp_dir = Path(os.environ.get("TEMP") or tempfile.gettempdir())
        dpath = Path(dflt["path"]) if dflt.get("path") else Path("")
        pattern_ok = dpath.name.startswith(f"hwkit_shot_{target_win['hwnd']}_") and dpath.suffix == ".png"
        check("默认输出路径落在 TEMP 且命名符合规范",
              dflt["ok"] and dpath.parent == temp_dir and pattern_ok,
              f"{dpath}")
        dpath.unlink(missing_ok=True)  # 清掉这条临时产物，别在用户 TEMP 里留垃圾

    nf = win.capture_window(title_contains="__绝对不存在的窗口标题__")
    check("找不到窗口时 ok=False 且有 error", nf["ok"] is False and bool(nf.get("error")),
          str(nf.get("error")))
    check("找不到窗口时 width/height/bytes 为 0",
          nf["width"] == 0 and nf["height"] == 0 and nf["bytes"] == 0)

    cam = win.capture_window(hwnd=0xDEADBEEF)
    check("无效 hwnd 时 ok=False", cam["ok"] is False, str(cam.get("error")))

    # 最小化分支：这是需求里明确要求"返回 ok=False 并给出原因"的路径
    mini = [w for w in windows if w["minimized"]]
    if mini:
        mres = win.capture_window(hwnd=mini[0]["hwnd"])
        check("最小化窗口 ok=False 且说明原因",
              mres["ok"] is False and "最小化" in str(mres.get("error", "")),
              f"{mres.get('error')} (hwnd={mini[0]['hwnd']} {mini[0]['title'][:24]!r})")
    else:
        print("    SKIP: 当前没有最小化窗口，该分支本次未覆盖")

    # visible_only=False 必须是可见集合的超集（隐藏窗口也要能列出来）
    all_windows = win.list_windows(visible_only=False)
    vis_hwnds = {w["hwnd"] for w in windows}
    all_hwnds = {w["hwnd"] for w in all_windows}
    check("visible_only=False 是可见集合的超集",
          vis_hwnds <= all_hwnds and len(all_hwnds) >= len(vis_hwnds),
          f"visible={len(vis_hwnds)} all={len(all_hwnds)}")

    print("\n[2c] find_window()")
    if target_win:
        f1 = win.find_window(pid=target_win["pid"])
        check("find_window 按 pid 命中", f1 is not None and f1["hwnd"] == target_win["hwnd"],
              f"hwnd={f1['hwnd'] if f1 else None}")
        f2 = win.find_window(title_contains=target_win["title"][:6].upper())  # 大小写不敏感
        check("find_window 按标题(大写)命中", f2 is not None and f2["hwnd"] == target_win["hwnd"],
              f"hwnd={f2['hwnd'] if f2 else None}")
    check("find_window 无匹配返回 None", win.find_window(pid=999_999_999) is None)

    # ---------------- 测试 2d: 离屏 / 被遮挡窗口 ----------------
    print("\n[2d] 离屏窗口抓取（证明不依赖屏幕像素/不受遮挡影响）")
    scr_w0 = int(_u32.GetSystemMetrics(0))
    scr_h0 = int(_u32.GetSystemMetrics(1))

    def _offscreen(w: dict) -> bool:
        r = w["rect"]
        return (
            r["width"] >= 200
            and r["height"] >= 150
            and (r["x"] + r["width"] <= 0 or r["y"] + r["height"] <= 0
                 or r["x"] >= scr_w0 or r["y"] >= scr_h0)
        )

    off_cands = [w for w in windows if not w["minimized"] and _offscreen(w)]
    if off_cands:
        off = off_cands[0]
        r = off["rect"]
        res = win.capture_window(hwnd=off["hwnd"], out_path=str(OUT_DIR / "shot_offscreen.png"))
        print(f"    离屏窗口 hwnd={off['hwnd']} rect=({r['x']},{r['y']} {r['width']}x{r['height']}) "
              f"title={off['title'][:30]!r}")
        if check("离屏窗口抓取 ok=True", res.get("ok") is True, res.get("error", "")):
            w3, h3, bpp3, px3 = decode_png(Path(res["path"]))
            c3 = distinct_colors(px3, bpp3)
            # 该窗口的矩形完全在屏幕外：BitBlt 屏幕裁剪拿不到任何东西，能抓到内容只可能来自
            # PrintWindow 让窗口自绘。这是"非侵入 + 不受遮挡"最硬的证据。
            check("离屏窗口抓到真实内容(>50 色)", c3 > 50, f"{c3} 种颜色, {w3}x{h3}")
    else:
        print("    SKIP: 当前桌面上没有完全离屏的窗口，该分支本次未覆盖")

    # ---------------- 测试 2e: 通道顺序 / 上下翻转正确性 ----------------
    # 这条是确定性的：手工填一块已知 BGRA 的 DIB，走 _dib_bits_to_rgb + _encode_png，
    # 再用本文件的解码器读回来逐字节比对。它能精确地盯住两类最难用肉眼发现、又最容易犯的
    # 编码 bug：BGRA→RGB 通道顺序搞错、DIB 自下而上的行序没翻转。
    print("\n[2e] 编码管线单元测试（BGRA→RGB 通道顺序 + 上下翻转 + PNG 往返）")
    probe = win._MemDib(4, 3)  # 4x3，够小，可以逐字节断言
    try:
        # DIB 是自下而上的：下面第 0 行其实是画面的最末一行。
        # 每行 4 个像素，每个像素 BGRA 四字节。故意让每行颜色不同，翻转错了一眼能看出来。
        dib_rows_bottom_up = [
            [(1, 2, 3, 255), (4, 5, 6, 255), (7, 8, 9, 255), (10, 11, 12, 255)],
            [(13, 14, 15, 255), (16, 17, 18, 255), (19, 20, 21, 255), (22, 23, 24, 255)],
            [(25, 26, 27, 255), (28, 29, 30, 255), (31, 32, 33, 255), (34, 35, 36, 255)],
        ]
        blob = bytearray()
        for row in dib_rows_bottom_up:
            for b, g, r, a in row:
                blob += bytes((b, g, r, a))
        ctypes.memmove(probe.bits, bytes(blob), len(blob))

        rgb = probe.rgb()
        # 期望：行序翻转（最后一行变第一行），通道 BGRA -> RGB
        expected = bytearray()
        for row in reversed(dib_rows_bottom_up):
            for b, g, r, _a in row:
                expected += bytes((r, g, b))
        check("_dib_bits_to_rgb 输出与期望逐字节一致", rgb == bytes(expected),
              f"len={len(rgb)}")

        png_bytes = win._encode_png(rgb, 4, 3)
        tmp_png = OUT_DIR / "unit_roundtrip.png"
        tmp_png.write_bytes(png_bytes)
        w_u, h_u, bpp_u, px_u = decode_png(tmp_png)
        check("PNG 往返尺寸正确", (w_u, h_u, bpp_u) == (4, 3, 3), f"{w_u}x{h_u} bpp={bpp_u}")
        check("PNG 往返像素与期望一致", px_u == bytes(expected))
        check("png_size() 读回正确尺寸", win.png_size(str(tmp_png)) == {"width": 4, "height": 3})
    finally:
        probe.close()

    bad = OUT_DIR / "unit_bad.png"
    bad.write_bytes(b"not a png at all, just text padding to be long enough....")
    try:
        win.png_size(str(bad))
        check("png_size 对非 PNG 抛 ValueError", False, "没有抛异常")
    except ValueError as exc:
        check("png_size 对非 PNG 抛 ValueError", True, str(exc)[:60])

    # ---------------- 测试 2f: 真实窗口的几何/通道交叉验证 ----------------
    # 用"不透明 + 无遮挡 + 完整在屏"的窗口做比对：此时 PrintWindow 结果与屏幕同区域位图应该
    # 几乎相同，能独立验证 DPI 映射、通道顺序与行翻转在真实数据上成立。
    # 为什么不用任务栏(Shell_TrayWnd)做这条：实测它的 PrintWindow 结果与屏幕裁剪差 90%，
    # 原因不是抓错，而是任务栏是亚克力/半透明表面 —— PrintWindow 渲染的是窗口本身（不带壁纸
    # 背衬），屏幕上看到的是它和壁纸混合后的颜色。内容完全一致（图标/文字都在），只是背景
    # 混色不同。所以这类窗口天然不适合做逐像素比对。
    print("\n[2f] 不透明无遮挡窗口 PrintWindow vs 屏幕同区域（真实数据几何校验）")

    def _intersects(ra: dict, rb: dict) -> bool:
        return not (
            ra["x"] + ra["width"] <= rb["x"]
            or rb["x"] + rb["width"] <= ra["x"]
            or ra["y"] + ra["height"] <= rb["y"]
            or rb["y"] + rb["height"] <= ra["y"]
        )

    _NON_OPAQUE = {"Shell_TrayWnd", "Progman", "WorkerW", "Windows.UI.Core.CoreWindow",
                   "XamlExplorerHostIslandWindow", "ForegroundStaging"}
    top_win = None
    for i, w in enumerate(windows):  # windows 顺序 = EnumWindows 顺序 = Z 序自上而下
        r = w["rect"]
        if w["minimized"] or w["class"] in _NON_OPAQUE:
            continue
        if r["width"] < 200 or r["height"] < 120:
            continue
        if not (r["x"] >= 0 and r["y"] >= 0
                and r["x"] + r["width"] <= scr_w0 and r["y"] + r["height"] <= scr_h0):
            continue
        if any(_intersects(o["rect"], r) for o in windows[:i]):
            continue  # 被更高 Z 序窗口压住，屏幕裁剪不可信
        top_win = w
        break

    if top_win is None:
        print("    SKIP: 桌面状态里没有同时满足[不透明/完整在屏/无遮挡]的窗口")
    else:
        r = top_win["rect"]
        pw = win.capture_window(hwnd=top_win["hwnd"], out_path=str(OUT_DIR / "topwin_pw.png"))
        sc = win.capture_screen(region=r, out_path=str(OUT_DIR / "topwin_screen.png"))
        print(f"    目标窗口 hwnd={top_win['hwnd']} class={top_win['class']} "
              f"{r['width']}x{r['height']} proc={top_win['process']} "
              f"title={top_win['title'][:30]!r}")
        if pw.get("ok") and sc.get("ok") and (pw["width"], pw["height"]) == (sc["width"], sc["height"]):
            wa, ha, bppa, pxa = decode_png(Path(pw["path"]))
            wb, hb, bppb, pxb = decode_png(Path(sc["path"]))
            n = wa * ha
            step = max(1, n // 200_000)
            diff = total = 0
            for i in range(0, n, step):
                oa, ob = i * bppa, i * bppb
                total += 1
                if (abs(pxa[oa] - pxb[ob]) > 8
                        or abs(pxa[oa + 1] - pxb[ob + 1]) > 8
                        or abs(pxa[oa + 2] - pxb[ob + 2]) > 8):
                    diff += 1
            ratio = diff / total if total else 1.0
            print(f"    逐像素差异率 = {ratio:.2%}（采样 {total} 像素）")
            check("PrintWindow 与屏幕同区域一致(<8%)", ratio < 0.08, f"{ratio:.2%}")
        else:
            print(f"    SKIP: 抓取失败或尺寸不一致 pw={pw.get('ok')} scr={sc.get('ok')}")

    # 遮挡场景的行为记录：这条不设断言，只把事实打出来——屏幕裁剪会被压住的上层窗口污染，
    # 而 PrintWindow 抓到的是完整内容，这正是本模块存在的理由。
    if target_win and shot:
        z = {w["hwnd"]: i for i, w in enumerate(windows)}
        idx = z.get(target_win["hwnd"], len(windows))
        occluders = [w for w in windows[:idx] if _intersects(w["rect"], target_win["rect"])]
        if occluders:
            names = ", ".join(f"{o['process']}({o['class']})" for o in occluders[:3])
            print(f"    [INFO] 被抓窗口被 {len(occluders)} 个更高 Z 序窗口遮挡: {names}")
            print("    [INFO] 该窗口的 PrintWindow 截图仍是完整未遮挡内容 —— 遮挡不影响抓取")

    # ---------------- 测试 3: capture_screen ----------------
    print("\n[3] capture_screen()")
    scr_w = int(_u32.GetSystemMetrics(0))
    scr_h = int(_u32.GetSystemMetrics(1))
    scr = win.capture_screen(out_path=str(OUT_DIR / "screen.png"))
    check("capture_screen ok=True", scr.get("ok") is True, scr.get("error", ""))
    if scr.get("ok"):
        check("屏幕尺寸 == getSystemMetrics",
              scr["width"] == scr_w and scr["height"] == scr_h,
              f"capture={scr['width']}x{scr['height']} metrics={scr_w}x{scr_h}")
        size = win.png_size(scr["path"])
        check("全屏 PNG 头尺寸与返回一致",
              size["width"] == scr["width"] and size["height"] == scr["height"], str(size))
        # 全屏图也要确认不是空白（只抽样）
        w2, h2, bpp2, px2 = decode_png(Path(scr["path"]))
        c2 = distinct_colors(px2, bpp2, max_samples=120_000)
        check("全屏图不是纯色", c2 > 50, f"{c2} 种颜色")
        print(f"    全屏 {size['width']}x{size['height']}, {scr['bytes'] / 1024 / 1024:.2f} MB, "
              f"{c2} 种颜色, {scr['bytes']} bytes")

    region = {"x": 100, "y": 100, "width": 640, "height": 360}
    reg = win.capture_screen(region=region, out_path=str(OUT_DIR / "region.png"))
    check("capture_screen(region) ok=True", reg.get("ok") is True, reg.get("error", ""))
    if reg.get("ok"):
        check("区域截图尺寸符合请求", reg["width"] == 640 and reg["height"] == 360,
              f"{reg['width']}x{reg['height']}")
    # 越界区域要被裁剪而不是失败/带黑边
    oob = win.capture_screen(region={"x": scr_w - 100, "y": scr_h - 80,
                                     "width": 500, "height": 500},
                             out_path=str(OUT_DIR / "region_oob.png"))
    check("越界区域被裁剪到屏幕内", oob.get("ok") is True and oob["width"] <= 100
          and oob["height"] <= 80, f"{oob.get('width')}x{oob.get('height')}")

    # ---------------- 测试 4: 非侵入性 ----------------
    print("\n[4] 非侵入性（不抢焦点 / 不动鼠标）")
    fg_after_hwnd = foreground()
    cursor_after = cursor_pos()
    windows_after = win.list_windows()
    fg_flags_after = {w["hwnd"]: w["foreground"] for w in windows_after}

    check("前台窗口句柄未变", fg_after_hwnd == fg_before_hwnd,
          f"{fg_before_hwnd} -> {fg_after_hwnd}")
    check("鼠标位置未变", cursor_after == cursor_before,
          f"{cursor_before} -> {cursor_after}")
    # 逐窗口比对 foreground 标记：这是"窗口被截图后有没有被提到前台"的直接证据
    changed = [
        h for h in fg_flags_before
        if h in fg_flags_after and fg_flags_before[h] != fg_flags_after[h]
    ]
    check("所有窗口的 foreground 标记未变", not changed, f"变化: {changed}")
    if target_win:
        check("被抓取窗口截前后都不是前台(说明没抢焦点)",
              fg_flags_before.get(target_win["hwnd"]) is False
              and fg_flags_after.get(target_win["hwnd"]) is False,
              f"before={fg_flags_before.get(target_win['hwnd'])} "
              f"after={fg_flags_after.get(target_win['hwnd'])}")

    # ---------------- 测试 5: GDI 句柄不泄漏 ----------------
    # MCP 服务器是长驻进程，反复截图如果漏 GDI 对象，跑几小时就会撞上 10000 句柄上限
    # （典型症状是 GDI 相关 API 开始失败）。这条直接量句柄增量。
    print("\n[5] GDI 句柄泄漏检查（30 次连续截图）")
    if target_win:
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.GetCurrentProcess.restype = wintypes.HANDLE
        _u32.GetGuiResources.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        _u32.GetGuiResources.restype = wintypes.DWORD
        proc_handle = k32.GetCurrentProcess()

        def gdi_objects() -> int:
            return int(_u32.GetGuiResources(proc_handle, 0))  # GR_GDIOBJECTS

        gdi_before = gdi_objects()
        n_ok = 0
        for _ in range(30):
            res = win.capture_window(hwnd=target_win["hwnd"],
                                     out_path=str(OUT_DIR / "leak_probe.png"))
            n_ok += 1 if res.get("ok") else 0
        gdi_after = gdi_objects()
        print(f"    成功 {n_ok}/30, GDI 对象 {gdi_before} -> {gdi_after} "
              f"(增量 {gdi_after - gdi_before})")
        # 注：这个进程本来就没有常驻 GDI 对象（实测裸 python 进程 GR_GDIOBJECTS=0），所以
        # 0 -> 0 不是"指标没生效"，而是"截图过程中创建的 DIB/DC 全部被 _MemDib 的 close 释放了"。
        check("30 次截图后 GDI 对象无显著增长(增量 < 10)",
              gdi_after - gdi_before < 10, f"增量 {gdi_after - gdi_before}")
    else:
        print("    SKIP: 没有可用目标窗口")

    # ---------------- 收尾 ----------------
    print("\n" + "=" * 78)
    print(f"通过 {len(_passed)} 项，失败 {len(_failed)} 项")
    if _failed:
        for f in _failed:
            print(f"  失败: {f}")
    print(f"截图产物目录: {OUT_DIR}")
    print("=" * 78)
    return 0 if not _failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
