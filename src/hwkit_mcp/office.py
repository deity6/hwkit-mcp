# -*- coding: utf-8 -*-
"""Office 文档转换（Word / Excel），走本机 Office 的 COM 接口。

为什么用 PowerShell 桥而不是 comtypes/pywin32：
本包刻意保持"除 MCP SDK 外零第三方依赖"，这样别人 clone 下来就能跑。
ctypes 直接调 IDispatch 也能做，但代码量是 PowerShell 桥的十倍且容易漏 Release；
本机既然装了 Office，`New-Object -ComObject Word.Application` 是最稳的一条路。

设计上的两个约束：
1. **不弹窗、不抢焦点**：Word/Excel 一律以 `Visible=$false`、`DisplayAlerts=0`
   打开，转换完立刻 Quit 并释放。用户正在用电脑时不会被抢到前台。
2. **不覆盖已有文件**：out 已存在时默认报错，除非显式 overwrite=True。
   作业交付物被静默覆盖是很难查的事故。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

# Word SaveAs 的格式常量（Word 的 WdSaveFormat）
WORD_FORMATS = {
    "docx": 16, "doc": 0, "pdf": 17, "odt": 23, "rtf": 6, "txt": 2,
}
# Excel SaveAs 的格式常量（XlFileFormat）
EXCEL_FORMATS = {
    "xlsx": 51, "xls": 56, "pdf": 57, "ods": 60, "csv": 6, "txt": 20,
}


def _pwsh(script: str, timeout: int = 180) -> tuple[int, str]:
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8") as f:
        f.write(script)
        path = f.name
    try:
        r = subprocess.run(
            ["pwsh", "-NoProfile", "-NonInteractive", "-File", path],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout,
        )
        return r.returncode, (r.stdout or "") + (r.stderr or "")
    except subprocess.TimeoutExpired:
        return 124, "pwsh 超时"
    finally:
        Path(path).unlink(missing_ok=True)


def _resolve_out(src: Path, fmt: str, out: str, overwrite: bool) -> Path:
    if out:
        dst = Path(out)
    else:
        dst = src.with_suffix("." + fmt)
    if dst.exists() and not overwrite:
        raise FileExistsError(f"目标已存在，拒绝覆盖：{dst}（需要覆盖请传 overwrite=True）")
    dst.parent.mkdir(parents=True, exist_ok=True)
    return dst


# ---------------------------------------------------------------- LibreOffice 兜底
def _soffice() -> str:
    """找 LibreOffice 的 soffice.exe。环境变量 HWKIT_SOFFICE 优先。"""
    p = os.environ.get("HWKIT_SOFFICE")
    if p and Path(p).exists():
        return p
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/usr/local/bin/soffice"):
        if Path(c).exists():
            return c
    return ""


def _soffice_convert(src: Path, fmt: str, dst: Path) -> dict:
    """用 LibreOffice 无头转换。COM 不可用时的兜底。

    注意 LibreOffice 只按"输出目录 + 由输入名推导的文件名"落盘，不能直接指定
    输出文件名，所以先转到临时目录再改名。
    """
    so = _soffice()
    if not so:
        return {"ok": False, "error": "既没有可用的 Office COM，也没装 LibreOffice"}
    tmpdir = Path(tempfile.mkdtemp(prefix="hwkit_lo_"))
    r = subprocess.run(
        [so, "--headless", "--norestore", "--convert-to", fmt,
         "--outdir", str(tmpdir), str(src)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240,
    )
    produced = list(tmpdir.glob(f"*.{fmt}"))
    if not produced:
        return {"ok": False, "engine": "libreoffice",
                "error": (r.stdout + r.stderr)[-1200:] or "LibreOffice 未产出文件"}
    shutil.move(str(produced[0]), str(dst))
    shutil.rmtree(tmpdir, ignore_errors=True)
    return {"ok": True, "engine": "libreoffice", "out": str(dst),
            "bytes": dst.stat().st_size, "error": None}


def com_probe() -> dict:
    """探测 Word / Excel 的 COM 自动化是否可用，并给出可执行建议。

    典型失败是 `CO_E_SERVER_EXEC_FAILURE (0x80080005)`：Word 进程本身能起，
    但 DCOM 激活被环境挡住。此时本模块会自动改走 LibreOffice（若已安装）。
    """
    ps = """
$ErrorActionPreference = 'Stop'
$out = [ordered]@{}
foreach ($pair in @(@('word','Word.Application'), @('excel','Excel.Application'))) {
    $key = $pair[0]; $prog = $pair[1]
    try {
        $app = New-Object -ComObject $prog
        $out[$key] = @{ ok = $true; version = $app.Version }
        try { $app.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    } catch {
        $out[$key] = @{ ok = $false; error = $_.Exception.Message }
    }
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
[Console]::OpenStandardOutput().Write(
    [Text.Encoding]::UTF8.GetBytes(($out | ConvertTo-Json -Compress -Depth 5)))
"""
    rc, txt = _pwsh(ps, timeout=120)
    try:
        start = txt.find("{")
        data = json.loads(txt[start:]) if start >= 0 else {}
    except Exception:  # noqa: BLE001
        data = {}
    so = _soffice()
    return {
        "word": data.get("word", {"ok": False, "error": txt[-600:]}),
        "excel": data.get("excel", {"ok": False, "error": txt[-600:]}),
        "libreoffice": {"path": so, "ok": bool(so)},
        "usable": bool(data.get("word", {}).get("ok") or so),
        "advice": ("COM 不可用时请设置环境变量 HWKIT_SOFFICE 指向 soffice.exe，"
                   "或改用 DSH 的 office_convert 工具。")
        if not data.get("word", {}).get("ok") else "COM 可用。",
    }


def word_convert(src: str, fmt: str, out: str = "", overwrite: bool = False) -> dict:
    """用 Word 把文档转成另一种格式（pdf / docx / doc / odt / rtf / txt）。

    典型用途：课程作业的 `.doc` 报告模板 → `.docx` 编辑 → `.pdf` 提交。

    Returns:
        {"ok": bool, "out": str, "bytes": int, "error": str|None}
    """
    fmt = fmt.lower().lstrip(".")
    if fmt not in WORD_FORMATS:
        return {"ok": False, "error": f"不支持的格式 {fmt}，可选：{sorted(WORD_FORMATS)}"}
    s = Path(src)
    if not s.exists():
        return {"ok": False, "error": f"源文件不存在：{src}"}
    try:
        dst = _resolve_out(s, fmt, out, overwrite)
    except FileExistsError as e:
        return {"ok": False, "error": str(e)}

    code = WORD_FORMATS[fmt]
    ps = f"""
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $doc = $word.Documents.Open("{str(s).replace(chr(92), chr(92)*2)}",
                                [ref]$false, [ref]$false, [ref]$false)
    $doc.SaveAs([ref]"{str(dst).replace(chr(92), chr(92)*2)}", [ref]{code})
    $doc.Close([ref]$false)
    $doc = $null
    Write-Output "OK"
}} catch {{
    Write-Output ("FAIL: " + $_.Exception.Message)
}} finally {{
    if ($doc) {{ try {{ $doc.Close([ref]$false) }} catch {{}} }}
    if ($word) {{ try {{ $word.Quit() }} catch {{}} }}
    if ($word) {{ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}}
"""
    rc, txt = _pwsh(ps)
    ok = "OK" in txt and dst.exists() and dst.stat().st_size > 0
    if ok:
        return {"ok": True, "engine": "word-com", "out": str(dst),
                "bytes": dst.stat().st_size, "error": None}
    # COM 激活被环境挡住时自动改走 LibreOffice，避免整条链路死掉
    lo = _soffice_convert(s, fmt, dst)
    if lo.get("ok"):
        return lo
    return {
        "ok": False, "engine": "none", "out": "",
        "bytes": 0,
        "error": (txt.strip()[-1200:] or f"rc={rc}") + " || " + str(lo.get("error")),
    }


def excel_convert(src: str, fmt: str, out: str = "", sheet: str = "",
                  overwrite: bool = False) -> dict:
    """用 Excel 转换工作簿（pdf / csv / xlsx / xls / ods）。

    `sheet` 只对 csv 有意义：多表工作簿必须指定表名，否则 Excel 只导出活动表。
    """
    fmt = fmt.lower().lstrip(".")
    if fmt not in EXCEL_FORMATS:
        return {"ok": False, "error": f"不支持的格式 {fmt}，可选：{sorted(EXCEL_FORMATS)}"}
    s = Path(src)
    if not s.exists():
        return {"ok": False, "error": f"源文件不存在：{src}"}
    try:
        dst = _resolve_out(s, fmt, out, overwrite)
    except FileExistsError as e:
        return {"ok": False, "error": str(e)}

    code = EXCEL_FORMATS[fmt]
    activate = ""
    if sheet and fmt == "csv":
        activate = (f'$wb.Worksheets.Item("{sheet}").Activate(); '
                    f'$wb.Application.ActiveSheet.Copy()')

    ps = f"""
$ErrorActionPreference = 'Stop'
$app = $null
$wb = $null
try {{
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Open("{str(s).replace(chr(92), chr(92)*2)}")
    {activate}
    $wb.SaveAs("{str(dst).replace(chr(92), chr(92)*2)}", {code})
    $wb.Close($false)
    $wb = $null
    if ($app.Workbooks.Count -gt 0) {{ $app.Workbooks.Item(1).Close($false) }}
    Write-Output "OK"
}} catch {{
    Write-Output ("FAIL: " + $_.Exception.Message)
}} finally {{
    if ($wb) {{ try {{ $wb.Close($false) }} catch {{}} }}
    if ($app) {{ try {{ $app.Quit() }} catch {{}} }}
    if ($app) {{ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}}
"""
    rc, txt = _pwsh(ps)
    ok = "OK" in txt and dst.exists() and dst.stat().st_size > 0
    if ok:
        return {"ok": True, "engine": "excel-com", "out": str(dst),
                "bytes": dst.stat().st_size, "error": None}
    # COM 激活被环境挡住时自动改走 LibreOffice，避免整条链路死掉
    lo = _soffice_convert(s, fmt, dst)
    if lo.get("ok"):
        return lo
    return {
        "ok": False, "engine": "none", "out": "",
        "bytes": 0,
        "error": (txt.strip()[-1200:] or f"rc={rc}") + " || " + str(lo.get("error")),
    }


def office_report(src: str) -> dict:
    """读 PDF 头部等基本信息，用来确认产物是合法 PDF（不解析内容）。"""
    p = Path(src)
    if not p.exists():
        return {"ok": False, "error": "文件不存在"}
    head = p.open("rb").read(8)
    return {
        "ok": True,
        "bytes": p.stat().st_size,
        "is_pdf": head.startswith(b"%PDF-"),
        "head": head.decode("latin-1", "replace"),
    }
