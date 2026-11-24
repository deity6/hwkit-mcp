"""用微软自家的 TMDL 解析器校验 PBIP 语义模型（不需要启动 Power BI Desktop）。

原理：Power BI Desktop 安装目录里带着 ``Microsoft.PowerBI.Tabular.dll``，其中的
``Microsoft.AnalysisServices.Tabular.TmdlSerializer.DeserializeDatabaseFromFolder()``
就是 Desktop 打开 PBIP 时用来把 ``definition/**.tmdl`` 反序列化成
``Database``/``Model``/``Table``/``Column``/``Measure`` 的那套解析器。这里把加载与调用
放进一个临时 ``.ps1``，用 ``pwsh -NoProfile -NonInteractive -File`` 跑，结果序列化成
JSON 写到 stdout —— 这样既不拖累 Python 进程（DLL 是 .NET 程序集），又能拿到
**比手写 TMDL 语法检查器可靠得多**的真解析结果。

要点（踩过的坑）
----------------
* **不要注册 ``AppDomain.AssemblyResolve`` 事件处理器**：之前这么干会触发 PowerShell
  自身的程序集解析路径，导致递归崩溃（栈溢出/进程被终止）。
  ``[Reflection.Assembly]::LoadFrom($dll)`` 单个文件即可 —— LoadFrom 上下文会自动
  在同一目录里探测依赖（``Microsoft.AnalysisServices.*.dll``），足够加载 Tabular.dll。
* 输出 JSON 时用 ``[Console]::OpenStandardOutput()`` 直接写 UTF-8 字节，
  不要用 ``Write-Output``/``Write-Host`` —— 否则 PowerShell 的序列化/格式化噪音会混进 stdout。
* 脚本文件本身存成带 BOM 的 UTF-8（PS 5.1 无 BOM 时按 ANSI 读，中文路径会乱）。
  脚本内部文本一律用纯 ASCII，避免编码问题；中文只出现在被校验的路径与 TMDL 内容里。
* 异常链：``TmdlSerializer`` 的格式错误可能是空的 ``TmdlFormatException``，
  也可能被 PowerShell 包成 ``MethodInvocationException``，所以 ``error`` 取**最内层**
  异常的 Message（根因，含 ``UnknownKeyword`` / 行号 / 出错行原文），
  ``detail`` 保留完整异常链（带类型名）便于定位。

主要 API：``validate_tmdl(definition_dir)`` / ``validate_pbip(pbip_path)``。
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

__all__ = [
    "validate_tmdl",
    "validate_pbip",
    "find_tabular_dll",
    "find_powershell",
    "TABULAR_DLL_NAME",
]

TABULAR_DLL_NAME = "Microsoft.PowerBI.Tabular.dll"

# Microsoft.PowerBI.Tabular.dll 的常见位置（Power BI Desktop 安装/Store 版）
_DLL_CANDIDATES = (
    r"C:\Program Files\Microsoft Power BI Desktop\bin",
    r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin",
    r"C:\Program Files\Microsoft Power BI Desktop RS\bin",
)

# --------------------------------------------------------------------------- #
# 核心 ps1：加载 Tabular.dll，反序列化 TMDL 目录，把结果 JSON 打到 stdout
# --------------------------------------------------------------------------- #
_PS1_SCRIPT = r"""param(
    [Parameter(Mandatory=$true)][string]$DefinitionDir,
    [Parameter(Mandatory=$true)][string]$TabularDll
)

$ErrorActionPreference    = 'Stop'
$ProgressPreference       = 'SilentlyContinue'
$WarningPreference        = 'SilentlyContinue'
$InformationPreference    = 'SilentlyContinue'
$ConfirmPreference        = 'None'

$detail = New-Object System.Collections.ArrayList
function Add-Detail([string]$msg) { [void]$detail.Add($msg) }

function Get-ExceptionChain([System.Exception]$ex) {
    $list = New-Object System.Collections.ArrayList
    $cur  = $ex
    $i    = 0
    while ($null -ne $cur -and $i -lt 8) {
        [void]$list.Add($cur.GetType().FullName + ': ' + $cur.Message)
        $cur = $cur.InnerException
        $i   = $i + 1
    }
    return ,$list
}

$result = [ordered]@{
    ok                  = $false
    compatibility_level = $null
    culture             = $null
    tables              = @()
    relationships       = @()
    error               = $null
    detail              = @()
}

function Write-ResultJson {
    $result.detail = @($detail.ToArray())
    $json = ConvertTo-Json -InputObject $result -Depth 8 -Compress
    $enc  = New-Object System.Text.UTF8Encoding($false)
    $out  = [Console]::OpenStandardOutput()
    $bytes = $enc.GetBytes($json)
    $out.Write($bytes, 0, $bytes.Length)
    $out.Flush()
}

try {
    if (-not (Test-Path -LiteralPath $DefinitionDir -PathType Container)) {
        throw [System.IO.DirectoryNotFoundException]::new("definition dir not found: $DefinitionDir")
    }
    if ([string]::IsNullOrWhiteSpace($TabularDll) -or -not (Test-Path -LiteralPath $TabularDll -PathType Leaf)) {
        throw [System.IO.FileNotFoundException]::new("Microsoft.PowerBI.Tabular.dll not found: $TabularDll")
    }

    # NOTE: intentionally NO AppDomain.AssemblyResolve handler here (it caused a
    # recursive PowerShell crash). LoadFrom probes dependencies in the same folder.
    $asm = [System.Reflection.Assembly]::LoadFrom($TabularDll)
    Add-Detail ('assembly=' + $asm.FullName)
    Add-Detail ('dll=' + $TabularDll)
    Add-Detail ('runtime=' + [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-Detail ('powershell=' + $PSVersionTable.PSVersion.ToString())

    $db = [Microsoft.AnalysisServices.Tabular.TmdlSerializer]::DeserializeDatabaseFromFolder($DefinitionDir)

    $result.ok = $true
    if (-not [string]::IsNullOrEmpty([string]$db.Name)) { Add-Detail ('database=' + [string]$db.Name) }
    if ($null -ne $db.CompatibilityLevel) { $result.compatibility_level = [int]$db.CompatibilityLevel }
    if ($null -ne $db.Model) {
        $culture = [string]$db.Model.Culture
        if (-not [string]::IsNullOrEmpty($culture)) { $result.culture = $culture }
    }

    $tables = New-Object System.Collections.ArrayList
    foreach ($t in @($db.Model.Tables)) {
        $row = [ordered]@{
            name       = [string]$t.Name
            columns    = [int]@($t.Columns).Count
            measures   = [int]@($t.Measures).Count
            partitions = [int]@($t.Partitions).Count
        }
        [void]$tables.Add($row)
    }
    if ($tables.Count -gt 0) { $result.tables = $tables.ToArray() } else { $result.tables = @() }

    $rels = New-Object System.Collections.ArrayList
    foreach ($r in @($db.Model.Relationships)) {
        $ft = [string]$r.FromTable.Name
        $fc = [string]$r.FromColumn.Name
        $tt = [string]$r.ToTable.Name
        $tc = [string]$r.ToColumn.Name
        $row = [ordered]@{
            from       = $ft + '[' + $fc + ']'
            to         = $tt + '[' + $tc + ']'
            fromTable  = $ft
            fromColumn = $fc
            toTable    = $tt
            toColumn   = $tc
        }
        [void]$rels.Add($row)
    }
    if ($rels.Count -gt 0) { $result.relationships = $rels.ToArray() } else { $result.relationships = @() }

    Add-Detail ('tables=' + @($db.Model.Tables).Count + ' relationships=' + @($db.Model.Relationships).Count)
}
catch {
    $chain = Get-ExceptionChain $_.Exception
    foreach ($line in $chain) { Add-Detail $line }
    if ($chain.Count -gt 0) {
        # 最内层 = 根因（例如 TmdlFormatException 的 UnknownKeyword 说明）
        $result.error = [string]$chain[$chain.Count - 1]
    } else {
        $result.error = 'unknown error'
    }
    $result.ok = $false
}

Write-ResultJson
"""


# --------------------------------------------------------------------------- #
# 环境探测
# --------------------------------------------------------------------------- #
def find_tabular_dll(bin_dir: str = "") -> str | None:
    """定位 ``Microsoft.PowerBI.Tabular.dll``。"""
    candidates: list[str] = []
    if bin_dir:
        candidates.append(os.path.join(bin_dir, TABULAR_DLL_NAME))
    env_bin = os.environ.get("POWERBI_BIN_DIR", "")
    if env_bin:
        candidates.append(os.path.join(env_bin, TABULAR_DLL_NAME))
    env_dll = os.environ.get("POWERBI_TABULAR_DLL", "")
    if env_dll:
        candidates.append(env_dll)
    candidates += [os.path.join(d, TABULAR_DLL_NAME) for d in _DLL_CANDIDATES]
    for cand in candidates:
        if cand and os.path.isfile(cand):
            return cand
    return None


def find_powershell() -> tuple[str, list[str]] | None:
    """返回 ``(exe, 前缀参数)``；优先 pwsh，退化到 Windows PowerShell 5.1。"""
    pwsh = shutil.which("pwsh")
    if pwsh:
        return pwsh, ["-NoProfile", "-NonInteractive"]
    legacy = shutil.which("powershell") or shutil.which("powershell.exe")
    if not legacy:
        for cand in (
            r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
            r"C:\Program Files\PowerShell\7\pwsh.exe",
        ):
            if os.path.isfile(cand):
                legacy = cand
                break
    if not legacy:
        return None
    return legacy, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass"]


def _empty_result(error: str = "", detail: list[str] | None = None) -> dict:
    return {
        "ok": False,
        "compatibility_level": None,
        "culture": None,
        "tables": [],
        "relationships": [],
        "error": error or None,
        "detail": detail or [],
    }


def _as_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    return [value]  # PowerShell 把单元素数组序列化成了对象


def _extract_json(stdout: str) -> dict | None:
    """从 stdout 里挖出 JSON（容忍前后噪音）。"""
    text = stdout.strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        try:
            return json.loads(text[start : end + 1])
        except json.JSONDecodeError:
            return None
    return None


def _normalize(payload: dict) -> dict:
    out = _empty_result()
    out["ok"] = bool(payload.get("ok"))
    cl = payload.get("compatibility_level")
    out["compatibility_level"] = int(cl) if isinstance(cl, (int, float)) else None
    culture = payload.get("culture")
    out["culture"] = str(culture) if culture else None
    tables = []
    for row in _as_list(payload.get("tables")):
        if not isinstance(row, dict):
            continue
        tables.append(
            {
                "name": str(row.get("name") or ""),
                "columns": int(row.get("columns") or 0),
                "measures": int(row.get("measures") or 0),
                "partitions": int(row.get("partitions") or 0),
            }
        )
    out["tables"] = tables
    rels = []
    for row in _as_list(payload.get("relationships")):
        if not isinstance(row, dict):
            continue
        rels.append(
            {
                "from": str(row.get("from") or ""),
                "to": str(row.get("to") or ""),
                "from_table": str(row.get("fromTable") or ""),
                "from_column": str(row.get("fromColumn") or ""),
                "to_table": str(row.get("toTable") or ""),
                "to_column": str(row.get("toColumn") or ""),
            }
        )
    out["relationships"] = rels
    err = payload.get("error")
    out["error"] = str(err) if err else None
    out["detail"] = [str(d) for d in _as_list(payload.get("detail"))]
    return out


# --------------------------------------------------------------------------- #
# 主 API
# --------------------------------------------------------------------------- #
def validate_tmdl(definition_dir: str, timeout: int = 180) -> dict:
    """用微软的 ``TmdlSerializer`` 反序列化 TMDL 目录。

    返回::

        {"ok", "compatibility_level", "culture",
         "tables": [{"name", "columns", "measures", "partitions"}, ...],
         "relationships": [{"from", "to", ...}, ...],
         "error", "detail": [...]}

    ``ok=False`` 时 ``error`` 是根因异常的 Message（例如
    ``TMDL Format Error: ... Parsing error type - UnknownKeyword ...``），
    ``detail`` 是完整异常链 + 环境信息。
    """
    src = Path(definition_dir)
    if not src.is_dir():
        return _empty_result(f"definition 目录不存在: {src}", [f"definition_dir={src}"])

    ps = find_powershell()
    if ps is None:
        return _empty_result(
            "找不到 PowerShell（pwsh / powershell.exe）", ["which(pwsh)=None"]
        )
    ps_exe, ps_args = ps

    dll = find_tabular_dll()
    if dll is None:
        return _empty_result(
            f"找不到 {TABULAR_DLL_NAME}（需要安装 Power BI Desktop）",
            ["checked: " + ", ".join(_DLL_CANDIDATES)],
        )

    tmpdir = tempfile.mkdtemp(prefix="hwkit_tmdl_")
    try:
        ps1 = Path(tmpdir) / "validate_tmdl.ps1"
        # 带 BOM 的 UTF-8：PS 5.1 无 BOM 时会按 ANSI 读脚本
        ps1.write_text(_PS1_SCRIPT, encoding="utf-8-sig")

        cmd = [ps_exe, *ps_args, "-File", str(ps1), str(src), str(dll)]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                timeout=timeout,
                text=False,
            )
        except subprocess.TimeoutExpired:
            return _empty_result(
                f"TMDL 解析超时（>{timeout}s）", [f"cmd={' '.join(cmd)}"]
            )

        stdout = proc.stdout.decode("utf-8", "replace")
        stderr = proc.stderr.decode("utf-8", "replace")
        payload = _extract_json(stdout)

        if payload is None:
            detail = [
                f"exit={proc.returncode}",
                f"dll={dll}",
                f"ps={ps_exe}",
                "stdout[:2000]=" + stdout[:2000],
                "stderr[:2000]=" + stderr[:2000],
            ]
            return _empty_result(
                "PowerShell 没有返回可解析的 JSON（见 detail）", detail
            )

        result = _normalize(payload)
        if stderr.strip():
            result["detail"].append("stderr=" + stderr.strip()[:1000])
        if proc.returncode != 0 and result.get("error") is None:
            result["ok"] = False
            result["error"] = f"PowerShell 退出码 {proc.returncode}"
        return result
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _read_json_file(path: Path) -> dict | None:
    for enc in ("utf-8-sig", "utf-16", "mbcs"):
        try:
            return json.loads(path.read_text(encoding=enc))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    return None


def validate_pbip(pbip_path: str, timeout: int = 180) -> dict:
    """从 ``.pbip`` 找到同名 ``.SemanticModel\\definition``，再调 :func:`validate_tmdl`。

    除了 ``validate_tmdl`` 的字段外，额外返回 ``pbip`` / ``report_path`` /
    ``model_dir`` / ``definition_dir``，方便调用方定位。
    """
    pbip = Path(pbip_path)
    if pbip.is_dir():
        found = sorted(pbip.glob("*.pbip"))
        if len(found) != 1:
            return _empty_result(
                f"目录里没有唯一的 .pbip（找到 {len(found)} 个）: {pbip}",
                [str(p) for p in found],
            )
        pbip = found[0]
    if not pbip.is_file():
        return _empty_result(f"pbip 不存在: {pbip}", [])

    base = pbip.parent
    candidates: list[Path] = []
    detail = [f"pbip={pbip}"]

    manifest = _read_json_file(pbip) or {}
    report_dir: Path | None = None
    for artifact in manifest.get("artifacts") or []:
        rel = ((artifact or {}).get("report") or {}).get("path")
        if rel:
            report_dir = (base / rel).resolve()
            detail.append(f"artifacts.report.path={rel}")
            break

    # 1) 报表的 definition.pbir 指向的模型目录（最权威）
    if report_dir is not None:
        pbir_def = report_dir / "definition.pbir"
        if pbir_def.is_file():
            pbir = _read_json_file(pbir_def) or {}
            by_path = ((pbir.get("datasetReference") or {}).get("byPath") or {}).get(
                "path"
            )
            if by_path:
                candidates.append((report_dir / by_path).resolve())
                detail.append(f"definition.pbir byPath={by_path}")
        # 2) 兄弟目录：X.Report -> X.SemanticModel
        if report_dir.name.lower().endswith(".report"):
            candidates.append(
                report_dir.with_name(report_dir.name[:-7] + ".SemanticModel")
            )

    # 3) pbip 同级的任意 *.SemanticModel
    for cand in sorted(base.glob("*.SemanticModel")):
        candidates.append(cand.resolve())

    seen: set[str] = set()
    for cand in candidates:
        key = str(cand).lower()
        if key in seen:
            continue
        seen.add(key)
        definition = cand / "definition"
        if (definition / "database.tmdl").is_file() or (definition / "model.tmdl").is_file():
            result = validate_tmdl(str(definition), timeout=timeout)
            result["pbip"] = str(pbip)
            result["report_path"] = str(report_dir) if report_dir else None
            result["model_dir"] = str(cand)
            result["definition_dir"] = str(definition)
            result["detail"] = detail + list(result.get("detail") or [])
            return result
        if (cand / "database.tmdl").is_file():  # 直接指着 definition 的目录
            result = validate_tmdl(str(cand), timeout=timeout)
            result["pbip"] = str(pbip)
            result["report_path"] = str(report_dir) if report_dir else None
            result["model_dir"] = str(cand)
            result["definition_dir"] = str(cand)
            result["detail"] = detail + list(result.get("detail") or [])
            return result

    result = _empty_result(
        f"找不到 SemanticModel 的 definition 目录（候选 {len(seen)} 个）", detail
    )
    result["pbip"] = str(pbip)
    result["report_path"] = str(report_dir) if report_dir else None
    result["model_dir"] = None
    result["definition_dir"] = None
    return result


if __name__ == "__main__":  # 手工排查用
    argv = sys.argv[1:]
    try:  # 输出一律 UTF-8，避免中文在 GBK 控制台/管道里变乱码
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    if not argv:
        print(
            "usage: python -m hwkit_mcp.modelcheck <definition_dir | X.pbip>",
            file=sys.stderr,
        )
        raise SystemExit(2)
    target = Path(argv[0])
    payload = (
        validate_pbip(str(target))
        if target.suffix.lower() == ".pbip"
        else validate_tmdl(str(target))
    )
    print(json.dumps(payload, ensure_ascii=False, indent=2))