# -*- coding: utf-8 -*-
"""hwkit-mcp —— 作业与 Power BI 辅助工具集的 MCP 服务器（stdio）。

设计立场
--------
本服务器只做两类事：**补缺口** 和 **不打扰**。

补缺口：Power BI 的 .pbix 打包、模型权威校验、报表离线渲染，这三件事在纯命令行
环境下没有顺手的工具；本服务器用 zip 手术 + 微软自家 TMDL 解析器 + DashCraft
渲染把它们补齐。

不打扰：所有工具都不得抢焦点、不得移动鼠标、不得发键盘事件。截图走
`PrintWindow(PW_RENDERFULLCONTENT)`，能抓到被遮挡甚至位于屏幕外的窗口内容，
而完全不改变窗口的前后层级。渲染走无头浏览器，不弹任何窗口。

关于第三方依赖
--------------
只有一条：MCP 官方 Python SDK（`mcp`）。其余全用标准库 + ctypes。
这样可以 `pip install .` 就在任何装了 Python 的机器上跑起来。

配套：pbix-mcp
--------------
.pbix 的**二进制语义**（建表、加度量值、改 ABF、DAX 求值）交给
[`pbix-mcp`](https://github.com/d0nk3yhm/pbix-mcp)，它是纯 Python 独立实现，
134 个工具，无需 Power BI Desktop。本服务器只负责它覆盖不到的"外壳"环节。
两者在 DSH 里作为两个 MCP 服务器并列注册。
"""
from __future__ import annotations

import json
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from . import modelcheck, office, pbipack, pbixbridge, render, win

mcp = FastMCP("hwkit")


# ============================================================ pbix-mcp 桥接
@mcp.tool()
def hw_pbix_tools() -> dict:
    """列出 pbix-mcp（配套的 Power BI 二进制工具箱）提供的全部工具名与一行说明。

    pbix-mcp 是纯 Python 独立实现的 .pbix 读写 + DAX 引擎，**不需要 Power BI Desktop**，
    能做本服务器不做的事：建表 / 加列 / 加度量值 / 改 ABF / 加视觉对象 / DAX 求值 /
    导出 TMDL 与 PBIP。它有 134 个工具，为避免全部 schema 常驻上下文，
    这里先用本工具查清单，再用 `hw_pbix_call` 按需调用。

    先看清单再调用：工具名和参数会随版本变化，别凭记忆猜。
    """
    return pbixbridge.pbix_tools()


@mcp.tool()
def hw_pbix_call(tool: str, arguments: dict | None = None) -> dict:
    """按需调用 pbix-mcp 的任意一个工具（134 个，先用 hw_pbix_tools 查名字）。

    例：
        hw_pbix_call("pbix_create", {"path": "...", "name": "..."})
        hw_pbix_call("pbix_evaluate_dax", {"path": "...", "expression": "SUM(T[x])"})

    返回 {"ok", "tool", "is_error", "result", "text"}；result 是工具的结构化返回，
    text 是原始文本（用于报错排查）。
    """
    return pbixbridge.pbix_call(tool, arguments)


# ============================================================ Power BI：文件层
@mcp.tool()
def hw_pbix_inspect(pbix_path: str) -> dict:
    """查看一个 .pbix / .pbit 的内部结构：条目清单、页面与视觉对象、内嵌模型大小、主题文件。

    用来在不动 Power BI Desktop 的前提下确认"包里到底有什么"，例如产出的 pbix 是否
    真的带了 DataModel、页面和图表有没有写进去。

    Args:
        pbix_path: .pbix / .pbit 文件路径。
    """
    return pbipack.inspect_pbix(pbix_path)


@mcp.tool()
def hw_pbix_extract(pbix_path: str, out_dir: str) -> dict:
    """把 .pbix 解包成目录（Report/、DataModel、[Content_Types].xml 等）。

    解包后可以用 Report/ 目录作为模板、做 diff、或直接改里面的 PBIR JSON。
    """
    return pbipack.extract_pbix(pbix_path, out_dir)


@mcp.tool()
def hw_pbix_build(report_dir: str, out_path: str, template_pbix: str,
                  extra_static_resources_dir: str = "") -> dict:
    """**无界面地把一个 PBIR 报表目录打包成 .pbix**。

    做法：拿一个已有的 .pbix 当外壳，保留它的 DataModel / Version /
    [Content_Types].xml / Settings / Metadata 等条目（连 ZipInfo 一起复用，
    保证 DataModel 仍是 stored(0) 不被重新压缩），只把 `Report/` 前缀下的内容
    换成 report_dir 里的最新文件。

    适用场景：模型没变、只是报表要改 —— 这是迭代看板时最常见的循环，
    整个过程不需要 Power BI Desktop。

    Args:
        report_dir: 形如 `X.Report` 的目录，里面应有 `definition/`（和可选的 `StaticResources/`）。
        out_path: 输出的 .pbix 路径。
        template_pbix: 提供 DataModel 的参考 pbix（模型结构必须与 report_dir 匹配）。
        extra_static_resources_dir: 追加的资源目录，内容并入 `Report/StaticResources/RegisteredResources/`。
    """
    return pbipack.build_pbix(report_dir, out_path, template_pbix, extra_static_resources_dir)


@mcp.tool()
def hw_pbix_pack(src_dir: str, out_path: str, datamodel_from: str = "") -> dict:
    """把 hw_pbix_extract 解出来的目录重新打回 .pbix（可用另一个文件的 DataModel 替换）。

    适合"先改解包目录里的 Report，再打回一个包"的流程。
    """
    return pbipack.pack_dir(src_dir, out_path, datamodel_from)


# ============================================================ Power BI：模型校验
@mcp.tool()
def hw_model_validate(target: str) -> dict:
    """用**微软自家的 TMDL 解析器**校验语义模型，等价于 Power BI Desktop 打开 PBIP
    时的模型解析步骤。

    为什么必须用它：其它开源 TMDL 解析器普遍比微软宽松。实测 DashCraft 会放过
    measure 下的 `description:` 属性，而微软的 TmdlSerializer 直接判定
    "Unsupported property" —— 于是出现"本地校验通过、Desktop 打不开"。
    本工具调的就是 Desktop 里那个 `Microsoft.PowerBI.Tabular.dll`。

    Args:
        target: `.pbip` 文件路径，或直接给 `...SemanticModel/definition` 目录。
    """
    return modelcheck.validate_pbip(target)


@mcp.tool()
def hw_model_validate_tmdl(definition_dir: str) -> dict:
    """同上，但直接接收 TMDL 的 `definition` 目录（省掉 pbip 解析）。"""
    return modelcheck.validate_tmdl(definition_dir)


# ============================================================ Power BI：渲染自查
@mcp.tool()
def hw_render_pages(source: str, out_dir: str, page: str = "", scale: int = 2,
                    keep_svg: bool = True) -> dict:
    """把报表**离线**渲染成每页一张 PNG，用于自查"做出来的看板长什么样"。

    不启动 Power BI Desktop，不弹窗，不抢焦点。链路是 DashCraft 出 SVG →
    统一系统字体栈并去掉错误裁切框 → 无头浏览器光栅化。

    除了图片，还会返回每页 SVG 里抽出的**全部可见文字** —— 这是核对图上到底写了
    什么最可靠的依据，不依赖图像识别。

    Args:
        source: `.pbip` / `.pbix` / `.pbit` / `*.Report` 目录。
        out_dir: PNG 输出目录。
        page: 只渲染某页（内部名/显示名/序号/显示名唯一子串）；空 = 全部页。
        scale: 每页单位对应像素数，2 表示 1280x720 的页面输出 2560x1440。
        keep_svg: 是否保留中间 SVG 便于排查。
    """
    return render.render_pages(source, out_dir, page, scale, keep_svg)


# ============================================================ Windows：非侵入式截图
@mcp.tool()
def hw_list_windows(visible_only: bool = True) -> list[dict]:
    """枚举桌面顶层窗口：hwnd、pid、标题、窗口类、位置尺寸、是否最小化、是否前台、进程名。

    先列窗口再决定抓哪个，比按标题猜更可靠。本工具纯读取，不改动任何窗口状态。
    """
    return win.list_windows(visible_only)


@mcp.tool()
def hw_capture_window(title_contains: str = "", pid: int = 0, hwnd: int = 0,
                      out_path: str = "") -> dict:
    """**非侵入式**抓取某个窗口的画面并存成 PNG。

    关键点：用 `PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT)` 直接让窗口自己把自己
    画到内存位图里，所以
      · 不调用 SetForegroundWindow，不改变窗口前后层级
      · 不移动鼠标、不发键盘
      · 窗口被别的窗口遮挡、甚至在屏幕外，照样能抓到完整内容
    这正是"让 agent 看一眼屏幕但不影响用户正常使用电脑"所需要的行为。

    Args:
        title_contains: 按标题子串匹配（大小写不敏感）。
        pid: 按进程 id 匹配（与 title_contains 二选一或并用）。
        hwnd: 直接指定窗口句柄（优先级最高）。
        out_path: 输出 PNG 路径；留空则写到临时目录。
    """
    return win.capture_window(title_contains, pid, out_path, hwnd)


@mcp.tool()
def hw_capture_screen(region: dict | None = None, out_path: str = "") -> dict:
    """抓取整个屏幕或指定区域。只读像素，不改动任何窗口状态。

    Args:
        region: 可选 `{"x","y","width","height"}`；超出屏幕范围会自动裁剪。
        out_path: 输出 PNG 路径；留空则写到临时目录。
    """
    return win.capture_screen(region, out_path)


# ============================================================ Office 文档转换
@mcp.tool()
def hw_com_probe() -> dict:
    """探测本机 Office COM 是否可用（Word / Excel 能否被脚本驱动）。

    转换类工具依赖 COM，若这里报 CO_E_SERVER_EXEC_FAILURE，说明本机 Office 的
    COM 激活被环境挡住了 —— 此时请改用 DSH 的 `office_convert` 工具，
    或装 LibreOffice 并设置环境变量 HWKIT_SOFFICE 指向 soffice.exe。
    """
    return office.com_probe()


@mcp.tool()
def hw_word_convert(src: str, fmt: str, out: str = "", overwrite: bool = False) -> dict:
    """用 Word 转换文档格式（pdf / docx / doc / odt / rtf / txt）。

    典型用途：课程作业给的 `.doc` 模板 → `.docx` 编辑 → `.pdf` 提交。
    转换过程 Word 以不可见方式运行，不抢焦点。目标文件已存在时默认拒绝覆盖。
    """
    return office.word_convert(src, fmt, out, overwrite)


@mcp.tool()
def hw_excel_convert(src: str, fmt: str, out: str = "", sheet: str = "",
                     overwrite: bool = False) -> dict:
    """用 Excel 转换工作簿（pdf / csv / xlsx / xls / ods）。

    `sheet` 只对 csv 有意义：多表工作簿不指定表名时 Excel 只导出活动表。
    """
    return office.excel_convert(src, fmt, out, sheet, overwrite)


@mcp.tool()
def hw_selfcheck() -> dict:
    """自检：报告各模块可用性与关键外部依赖（dashcraft / 无头浏览器 / TMDL 校验器）状态。

    部署到新机器后先跑这个，能立刻看出哪一环缺件。
    """
    out: dict = {"ok": True, "modules": {}, "deps": {}, "python": None}
    import sys
    out["python"] = sys.version.split()[0]

    out["modules"]["win"] = {"available": True,
                             "functions": [n for n in dir(win) if not n.startswith("_")]}
    out["modules"]["pbipack"] = {"available": True}
    out["modules"]["modelcheck"] = {"available": True}
    out["modules"]["render"] = {"available": True}
    out["modules"]["office"] = {"available": True}

    dc = render._dashcraft()
    out["deps"]["dashcraft"] = {"path": dc, "ok": bool(dc)}
    br = render._browser()
    out["deps"]["headless_browser"] = {"path": br, "ok": bool(br)}

    try:
        probe = office.com_probe()
        out["deps"]["office_com"] = {"ok": bool(probe.get("usable")), **probe}
    except Exception as e:  # noqa: BLE001
        out["deps"]["office_com"] = {"ok": False, "error": str(e)}

    try:
        tables_dll = Path(r"C:\Program Files\Microsoft Power BI Desktop\bin"
                          r"\Microsoft.PowerBI.Tabular.dll")
        out["deps"]["tmdl_validator"] = {
            "path": str(tables_dll), "ok": tables_dll.exists(),
            "note": "微软 TMDL 解析器；缺失时 hw_model_validate 不可用",
        }
    except Exception as e:  # noqa: BLE001
        out["deps"]["tmdl_validator"] = {"ok": False, "error": str(e)}

    out["ok"] = all(v.get("ok") for k, v in out["deps"].items()
                    if k in ("dashcraft", "headless_browser", "tmdl_validator"))
    return out


def main() -> None:
    """stdio 入口。"""
    mcp.run("stdio")


if __name__ == "__main__":
    main()


def _dump(obj) -> str:
    """调试辅助：把结果打成可读 JSON（不参与 MCP 协议）。"""
    return json.dumps(obj, ensure_ascii=False, indent=2, default=str)
