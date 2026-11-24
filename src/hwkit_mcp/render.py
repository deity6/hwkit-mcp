# -*- coding: utf-8 -*-
"""Power BI 报表离线渲染成 PNG —— 用于"看一眼自己做的看板长什么样"。

为什么需要这一层而不是直接调 DashCraft：

1. DashCraft 能把 PBIP / PBIX / PBIT 离线渲染成 SVG 或 PNG，但**不依赖 Power BI
   Desktop**，所以不会在用户桌面上弹窗、不抢焦点 —— 这正是本工具的前提。
2. 但它自带的字体是编译进二进制的 Selawik，既没有中文字形、也不在本机安装。
   直接出 PNG 时中文全是缺失告警、标题一片空白。
3. 出 SVG 再交给 Edge 无头光栅化可以解决中文字形（走系统字体回退），但数字仍按
   Selawik 的字体度量排版，换成别的字体就会出现裁切 / 错位。
4. 它给 cardVisual 的数字和标签套了一个约 25.7px 高的 clipPath，而数字字号是 60px，
   必然被腰斩 —— 这是它渲染器的布局缺陷，真机 Power BI 正常。

所以本模块做三件事：`dashcraft render --svg` → 后处理 SVG（统一系统字体栈、去掉
子裁切框）→ `msedge --headless --screenshot`。同时把 SVG 里所有可见文字抽出来返回，
这是核对"图上到底写了什么"最可靠的依据（不依赖图像识别）。
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

# ---------------------------------------------------------------- 可配置路径
# DashCraft 的查找顺序：环境变量 → PATH → 若干常见安装位置。
# 不写死绝对路径，这样项目 clone 到任何机器上都能用。
DASHCRAFT_ENV = "HWKIT_DASHCRAFT"
DASHCRAFT_NAMES = ["dashcraft.exe", "dashcraft"]
DASHCRAFT_HINTS = [
    Path.home() / "dashcraft",
    Path(os.environ.get("LOCALAPPDATA", "")) / "DashCraft",
    Path(os.environ.get("ProgramFiles", "")) / "DashCraft",
    Path.cwd() / "dashcraft",
]
EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "/usr/bin/microsoft-edge",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
]
FONT_STACK = "Microsoft YaHei, DengXian, Segoe UI, Arial, sans-serif"


def _dashcraft() -> str:
    """定位 dashcraft 可执行文件。"""
    env = os.environ.get(DASHCRAFT_ENV)
    if env and Path(env).exists():
        return env
    for name in DASHCRAFT_NAMES:
        hit = shutil.which(name)
        if hit:
            return hit
    for base in DASHCRAFT_HINTS:
        if not base or not str(base):
            continue
        try:
            for p in base.rglob(DASHCRAFT_NAMES[0]):
                return str(p)
        except Exception:  # noqa: BLE001
            continue
    return ""


def _browser() -> str:
    p = os.environ.get("HWKIT_BROWSER")
    if p and Path(p).exists():
        return p
    for c in EDGE_CANDIDATES:
        if Path(c).exists():
            return c
    return ""


def _run(cmd: list[str], timeout: int = 180) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=timeout,
    )


def _postprocess_svg(svg: str) -> str:
    """把 DashCraft 的 SVG 调成"用系统字体重排也说得通"的样子。"""
    # 1) 统一字体栈：Selawik / 'Segoe UI' / sans-serif → 系统中文字体优先
    svg = re.sub(r'font-family="[^"]*"', f'font-family="{FONT_STACK}"', svg)
    # 2) 去掉所有子裁切框（保留页面级 c0）。
    #    这些框是按 Selawik 的行高度量的，换字体 + cardVisual 数字 60px 时必然裁字。
    svg = re.sub(r'clip-path="url\(#c(?!0\))[^)]*\)"', "", svg)
    return svg


def _svg_texts(svg: str) -> list[str]:
    out = []
    for raw in re.findall(r"<text[^>]*>(.*?)</text>", svg, re.S):
        t = re.sub(r"&apos;", "'", raw)
        t = re.sub(r"&quot;", '"', t)
        t = re.sub(r"&amp;", "&", t)
        t = re.sub(r"&lt;", "<", t)
        t = re.sub(r"&gt;", ">", t)
        t = t.strip()
        if t:
            out.append(t)
    return out


def render_pages(
    source: str,
    out_dir: str,
    page: str = "",
    scale: int = 2,
    keep_svg: bool = True,
) -> dict:
    """把报表渲染成每页一张 PNG。

    Args:
        source: `.pbip` / `.pbix` / `.pbit` / `*.Report` 目录路径。
        out_dir: 输出目录（PNG 写这里）。
        page: 只渲染某页（内部名 / 显示名 / 序号 / 显示名唯一子串）。空=全部页。
        scale: 每页单位对应多少像素，2 表示 1280x720 的页面出 2560x1440。
        keep_svg: 是否保留中间 SVG（排查问题有用）。

    Returns:
        {"ok": bool, "pages": [{"index","name","png","width","height","bytes","texts":[...]}],
         "error": str|None, "warnings": [...]}
    """
    dc = _dashcraft()
    if not dc:
        return {"ok": False, "error": "找不到 dashcraft.exe，可用环境变量 HWKIT_DASHCRAFT 指定",
                "pages": []}
    src = Path(source)
    if not src.exists():
        return {"ok": False, "error": f"路径不存在: {source}", "pages": []}

    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    svg_dir = out / "_svg"
    if svg_dir.exists():
        shutil.rmtree(svg_dir, ignore_errors=True)
    svg_dir.mkdir(parents=True, exist_ok=True)

    cmd = [dc, "render", str(src), "--out-dir", str(svg_dir), "--format", "svg",
           "--force", "--scale", str(scale)]
    cmd += ["--page", page] if page else ["--all-pages"]

    try:
        r = _run(cmd)
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "dashcraft 渲染超时", "pages": []}

    warnings = [ln.strip() for ln in r.stdout.splitlines() if "WARN No fonts" in ln]
    if r.returncode != 0:
        return {"ok": False, "error": (r.stdout[-2000:] + r.stderr[-2000:]).strip(),
                "pages": [], "warnings": warnings}

    browser = _browser()
    svgs = sorted(p for p in svg_dir.glob("*.svg") if not p.name.startswith("fixed_"))
    if not svgs:
        return {"ok": False, "error": "dashcraft 没有产出 SVG", "pages": [],
                "warnings": warnings}

    pages = []
    for idx, sp in enumerate(svgs, 1):
        svg_raw = sp.read_text(encoding="utf-8")
        svg = _postprocess_svg(svg_raw)
        fixed = svg_dir / f"fixed_{sp.name}"
        fixed.write_text(svg, encoding="utf-8")

        m = re.search(r'<svg[^>]*width="([\d.]+)"[^>]*height="([\d.]+)"', svg)
        w, h = (int(float(m.group(1))), int(float(m.group(2)))) if m else (1280, 720)

        png = out / (sp.stem + ".png")
        if browser:
            rr = _run([browser, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                       f"--force-device-scale-factor={scale}",
                       f"--window-size={w},{h}",
                       f"--screenshot={png}",
                       "file:///" + fixed.as_posix()], timeout=120)
            ok = png.exists() and png.stat().st_size > 0
            if not ok:
                pages.append({"index": idx, "name": sp.stem, "png": "",
                              "error": (rr.stderr or "")[-800:], "texts": []})
                continue
        else:
            png = Path("")
            ok = False

        pages.append({
            "index": idx,
            "name": sp.stem,
            "png": str(png) if ok else "",
            "width": w, "height": h,
            "bytes": png.stat().st_size if ok else 0,
            "visual_count": svg.count('class="pbi-visual"'),
            "texts": _svg_texts(svg),
            "error": None if ok else "找不到可用浏览器（Edge/Chrome），可用 HWKIT_BROWSER 指定",
        })

    if not keep_svg:
        shutil.rmtree(svg_dir, ignore_errors=True)

    return {
        "ok": all(p.get("png") for p in pages) if pages else False,
        "pages": pages,
        "warnings": warnings[:5],
        "error": None,
    }