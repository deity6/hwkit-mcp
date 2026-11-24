"""PBIX 无界面打包 / 解包（纯标准库，零外部依赖）。

一个 ``.pbix`` 就是一个 ZIP 包，内部条目形如::

    Version
    [Content_Types].xml
    DiagramLayout
    Settings
    Metadata
    SecurityBindings
    DataModel                       # VertiPaq ABF，**stored(compress_type=0)，不压缩**
    Report/definition/**            # PBIR 报表定义，与磁盘上的 *.Report 目录同构
    Report/StaticResources/RegisteredResources/*.json   # 自定义主题

本模块的目标能力：**把磁盘上的 PBIR 报表目录换进一个已有的 .pbix 模板里，
产出新的 .pbix**，全程不启动 Power BI Desktop。

关键约定
--------
* 模板里所有非 ``Report/`` 前缀的条目（DataModel / Version / Settings / ...）
  原样搬运，并且**复用原始 ``ZipInfo``**（保住 compress_type，尤其是 DataModel 的
  stored(0)；一旦被 deflate 重新压缩，字节就不再逐位相同）。
* ``Report/`` 前缀的旧条目全部丢弃，再按下面的映射写入新目录内容::

      <src>/definition/**        -> Report/definition/**
      <src>/StaticResources/**   -> Report/StaticResources/**

  路径分隔符统一成 ``/``，条目名带 ``Report/`` 前缀。
* ``[Content_Types].xml`` 按需修补：只在新报表引入了模板里没有的扩展名时才插入
  ``<Default Extension="..." ContentType="..."/>``；不需要修补时原条目字节不动。

主要 API：``list_entries`` / ``inspect_pbix`` / ``build_pbix`` / ``extract_pbix`` / ``pack_dir``。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import time
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

__all__ = [
    "REPORT_PREFIX",
    "CONTENT_TYPES_ENTRY",
    "DATAMODEL_ENTRY",
    "list_entries",
    "inspect_pbix",
    "build_pbix",
    "extract_pbix",
    "pack_dir",
]

REPORT_PREFIX = "Report/"
CONTENT_TYPES_ENTRY = "[Content_Types].xml"
DATAMODEL_ENTRY = "DataModel"
DATAMODEL_SCHEMA_ENTRY = "DataModelSchema"  # .pbit 用的是 JSON 版模型

# OPC 里常见扩展名 -> ContentType。模板里缺哪个就补哪个。
_KNOWN_CONTENT_TYPES = {
    "json": "application/json",
    "xml": "application/xml",
    "pbir": "application/json",
    "pbism": "application/json",
    "tmdl": "text/plain",
    "txt": "text/plain",
    "csv": "text/csv",
    "html": "text/html",
    "htm": "text/html",
    "css": "text/css",
    "js": "application/javascript",
    "svg": "image/svg+xml",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "ico": "image/x-icon",
    "woff": "font/woff",
    "woff2": "font/woff2",
    "ttf": "font/ttf",
    "eot": "application/vnd.ms-fontobject",
    "wav": "audio/wav",
    "mp3": "audio/mpeg",
    "mp4": "video/mp4",
    "abf": "application/octet-stream",
    "bin": "application/octet-stream",
}
_FALLBACK_CONTENT_TYPE = "application/octet-stream"

# 通常以 stored(0) 形式落盘的二进制条目（媒体资源同样不该被重压缩）
_STORED_EXT_HINTS = {
    "png", "jpg", "jpeg", "gif", "bmp", "ico", "woff", "woff2", "ttf", "eot",
    "wav", "mp3", "mp4", "abf", "bin",
}

# 装饰性 visual 容器：不算"数据视觉"（inspect 的 visuals 计数会排除它们）
_DECORATIVE_VISUAL_TYPES = {
    "textbox",
    "shape",
    "basicshape",
    "image",
    "actionbutton",
    "pagenavigator",
    "bookmarknavigator",
    "qnavisual",
}


# --------------------------------------------------------------------------- #
# 小工具
# --------------------------------------------------------------------------- #
def _norm(name: str) -> str:
    """ZIP 条目名统一成 ``/`` 分隔。"""
    return str(name).replace("\\", "/").lstrip("/")


def _is_report_entry(name: str) -> bool:
    n = _norm(name)
    return n == "Report" or n.startswith(REPORT_PREFIX)


def decode_entry_bytes(raw: bytes) -> str:
    """把 PBIX 里的文本条目解成 str。

    PBIX 内部编码并不统一：``Version``/``Settings``/``Metadata`` 是 UTF-16LE（无 BOM），
    ``[Content_Types].xml`` 是带 BOM 的 UTF-8，PBIR 的 json 是 UTF-8。这里做启发式判断。
    """
    if not raw:
        return ""
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        try:
            return raw.decode("utf-16").lstrip("\ufeff")
        except UnicodeDecodeError:
            pass
    if raw[:3] == b"\xef\xbb\xbf":
        return raw[3:].decode("utf-8", "replace")
    if raw[:64].count(b"\x00") >= 2:  # 无 BOM 的 UTF-16LE：ASCII 字符之间夹 NUL
        try:
            return raw.decode("utf-16-le").replace("\x00", "")
        except UnicodeDecodeError:
            pass
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            return raw.decode("mbcs")
        except (UnicodeDecodeError, LookupError):
            return raw.decode("latin-1")


def _sha256(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _read_zip_entry_bytes(zf: zipfile.ZipFile, name: str) -> bytes | None:
    try:
        return zf.read(name)
    except KeyError:
        return None


# --------------------------------------------------------------------------- #
# 1) 列条目
# --------------------------------------------------------------------------- #
def list_entries(pbix_path: str) -> list[dict]:
    """列出 zip 条目。

    返回 ``[{"name", "size", "compressed", "compress_type"}, ...]``；
    ``compress_type`` 顺带给出（0=stored，8=deflate），用于确认 DataModel 没被重压缩。
    """
    path = Path(pbix_path)
    if not path.is_file():
        raise FileNotFoundError(f"pbix 不存在: {path}")
    with zipfile.ZipFile(path) as zf:
        return [
            {
                "name": info.filename,
                "size": int(info.file_size),
                "compressed": int(info.compress_size),
                "compress_type": int(info.compress_type),
            }
            for info in zf.infolist()
        ]


# --------------------------------------------------------------------------- #
# 2) 检查
# --------------------------------------------------------------------------- #
def _pages_from_pbir(zf: zipfile.ZipFile, names: list[str]) -> list[dict]:
    """从 ``Report/definition/pages/**`` 读出页面与 visual 数量（PBIR 布局）。

    ``visuals`` 统计的是**数据视觉**：每个页面的 ``hdr`` 之类 ``textbox`` / ``shape`` /
    ``image`` / 按钮 / 导航器属于装饰容器，不计入（真实样本每页 11 个 visual.json，
    其中 1 个是 ``hdr`` 文本框，所以数据视觉是 10 个）。
    装饰容器的数量与类型分布另见 ``visual_containers`` / ``visual_types``。
    """
    page_json_re = re.compile(r"^Report/definition/pages/([^/]+)/page\.json$", re.I)
    visual_re = re.compile(
        r"^Report/definition/pages/([^/]+)/visuals/.*visual\.json$", re.I
    )

    page_dirs: list[str] = []
    for name in names:
        m = page_json_re.match(name)
        if m:
            page_dirs.append(m.group(1))

    order: list[str] = []
    raw_pages_json = _read_zip_entry_bytes(zf, "Report/definition/pages/pages.json")
    if raw_pages_json is not None:
        try:
            meta = json.loads(decode_entry_bytes(raw_pages_json))
            order = [str(p) for p in (meta.get("pageOrder") or [])]
        except Exception:  # noqa: BLE001 - 坏 json 不该让 inspect 整体失败
            order = []

    known = set(page_dirs)
    ordered: list[str] = []
    for d in order:
        if d in known and d not in ordered:
            ordered.append(d)
    for d in sorted(page_dirs):
        if d not in ordered:
            ordered.append(d)

    visual_files: dict[str, list[str]] = {d: [] for d in ordered}
    for name in names:
        m = visual_re.match(name)
        if m and m.group(1) in visual_files:
            visual_files[m.group(1)].append(name)

    pages: list[dict] = []
    for page_dir in ordered:
        raw = _read_zip_entry_bytes(
            zf, f"Report/definition/pages/{page_dir}/page.json"
        )
        name = page_dir
        display = page_dir
        if raw is not None:
            try:
                meta = json.loads(decode_entry_bytes(raw))
                name = str(meta.get("name") or page_dir)
                display = str(meta.get("displayName") or name)
            except Exception:  # noqa: BLE001
                pass

        types: dict[str, int] = {}
        data_visuals = 0
        containers = 0
        for entry in sorted(visual_files.get(page_dir, [])):
            containers += 1
            visual_type = ""
            raw_visual = _read_zip_entry_bytes(zf, entry)
            if raw_visual is not None:
                try:
                    payload = json.loads(decode_entry_bytes(raw_visual))
                    visual_type = str(((payload.get("visual") or {}).get("visualType") or ""))
                except Exception:  # noqa: BLE001
                    visual_type = ""
            types[visual_type or "unknown"] = types.get(visual_type or "unknown", 0) + 1
            if (visual_type or "").lower() not in _DECORATIVE_VISUAL_TYPES:
                data_visuals += 1

        pages.append(
            {
                "name": name,
                "displayName": display,
                "visuals": data_visuals,
                "visual_containers": containers,
                "visual_types": types,
                "dir": page_dir,
            }
        )
    return pages


def _pages_from_legacy_layout(zf: zipfile.ZipFile) -> list[dict]:
    """老式 pbix（非 PBIR）布局兜底：``Report/Layout`` 是 UTF-16LE 的 JSON。

    老布局拿不到可靠的 visualType，这里一律按数据视觉计数（容器数即 visuals）。
    """
    raw = _read_zip_entry_bytes(zf, "Report/Layout")
    if raw is None:
        return []
    try:
        layout = json.loads(decode_entry_bytes(raw).lstrip("\ufeff"))
    except Exception:  # noqa: BLE001
        return []
    pages: list[dict] = []
    for section in layout.get("sections") or []:
        try:
            cfg = json.loads(section.get("config") or "{}")
        except Exception:  # noqa: BLE001
            cfg = {}
        count = len(section.get("visualContainers") or [])
        pages.append(
            {
                "name": str(section.get("name") or cfg.get("name") or ""),
                "displayName": str(
                    section.get("displayName") or cfg.get("displayName") or ""
                ),
                "visuals": count,
                "visual_containers": count,
                "visual_types": {},
            }
        )
    return pages


def inspect_pbix(pbix_path: str) -> dict:
    """扫描一个 .pbix / .pbit 的结构。

    返回::

        {"is_pbix", "is_pbit", "entries": [名字...],
         "pages": [{"name", "displayName", "visuals"}, ...],
         "has_datamodel", "datamodel_bytes", "theme_files": [...], "version"}

    ``pages[*].visuals`` 是**数据视觉**数量（排除 textbox/shape/image/按钮/导航器这类
    装饰容器；真实样本每页 11 个 visual.json，其中 1 个是 ``hdr`` 文本框 → 10）。
    另附 ``visual_containers``（容器总数）、``visual_types``（类型分布）、
    ``active_page`` / ``entry_count`` / ``datamodel_compress_type`` 等诊断字段。
    """
    path = Path(pbix_path)
    if not path.is_file():
        raise FileNotFoundError(f"pbix 不存在: {path}")

    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        names = [i.filename for i in infos]
        by_name = {i.filename: i for i in infos}

        dm_info = by_name.get(DATAMODEL_ENTRY)
        has_schema = DATAMODEL_SCHEMA_ENTRY in by_name
        is_pbit = has_schema and dm_info is None
        is_pbix = dm_info is not None or (
            not is_pbit and CONTENT_TYPES_ENTRY in by_name
        )

        pages = _pages_from_pbir(zf, names)
        if not pages:
            pages = _pages_from_legacy_layout(zf)

        active_page = None
        raw_pages_json = _read_zip_entry_bytes(zf, "Report/definition/pages/pages.json")
        if raw_pages_json is not None:
            try:
                active_page = json.loads(decode_entry_bytes(raw_pages_json)).get(
                    "activePageName"
                )
            except Exception:  # noqa: BLE001
                active_page = None

        version = ""
        raw_version = _read_zip_entry_bytes(zf, "Version")
        if raw_version is not None:
            version = decode_entry_bytes(raw_version).strip()

        theme_prefix = "Report/StaticResources/RegisteredResources/"
        theme_files = [
            n for n in names if _norm(n).startswith(theme_prefix) and not n.endswith("/")
        ]

        return {
            "is_pbix": bool(is_pbix),
            "is_pbit": bool(is_pbit),
            "entries": names,
            "entry_count": len(names),
            "pages": [
                {
                    "name": p["name"],
                    "displayName": p["displayName"],
                    "visuals": p["visuals"],
                    "visual_containers": p.get("visual_containers", p["visuals"]),
                    "visual_types": p.get("visual_types", {}),
                }
                for p in pages
            ],
            "has_datamodel": dm_info is not None,
            "datamodel_bytes": int(dm_info.file_size) if dm_info is not None else 0,
            "datamodel_compress_type": int(dm_info.compress_type)
            if dm_info is not None
            else None,
            "theme_files": theme_files,
            "version": version,
            "active_page": active_page,
            "path": str(path),
        }


# --------------------------------------------------------------------------- #
# 3) [Content_Types].xml 修补
# --------------------------------------------------------------------------- #
def _ensure_content_types(raw_ct: bytes, entry_names: list[str]) -> bytes | None:
    """确保 OPC 内容类型清单覆盖新条目用到的扩展名。

    只在真的缺 ``<Default Extension=...>`` 时改写并返回新字节；
    不需要修补时返回 ``None``（调用方据此原样搬运旧条目）。
    """
    if raw_ct is None:
        return None
    text = decode_entry_bytes(raw_ct)
    if not text.strip():
        return None

    existing = {
        m.group(1).lower()
        for m in re.finditer(
            r"""<Default\b[^>]*\bExtension\s*=\s*["']([^"']+)["']""", text, re.I
        )
    }
    need: set[str] = set()
    for name in entry_names:
        norm = _norm(name)
        if norm.endswith("/"):
            continue
        # 清单自身不是 OPC part，不该为它补一个 "xml" Default
        if norm == CONTENT_TYPES_ENTRY:
            continue
        suffix = PurePosixPath(norm).suffix
        if suffix:
            ext = suffix.lstrip(".").lower()
            if ext:
                need.add(ext)

    missing = sorted(ext for ext in need if ext not in existing)
    if not missing:
        return None

    added = "".join(
        '<Default Extension="%s" ContentType="%s"/>'
        % (ext, _KNOWN_CONTENT_TYPES.get(ext, _FALLBACK_CONTENT_TYPE))
        for ext in missing
    )
    idx = text.lower().rfind("</types>")
    if idx < 0:  # 没有根闭合标签，别瞎改
        return None
    patched_text = text[:idx] + added + text[idx:]

    # 保留原编码风格（BOM 有无 / UTF-16）
    if raw_ct[:3] == b"\xef\xbb\xbf":
        patched = b"\xef\xbb\xbf" + patched_text.encode("utf-8")
    elif raw_ct[:2] in (b"\xff\xfe", b"\xfe\xff"):
        patched = b"\xff\xfe" + patched_text.replace("\ufeff", "").encode("utf-16-le")
    else:
        patched = patched_text.encode("utf-8")

    ET.fromstring(decode_entry_bytes(patched).encode("utf-8"))  # 必须仍是良构 XML
    return patched


# --------------------------------------------------------------------------- #
# 4) 打包
# --------------------------------------------------------------------------- #
def _collect_report_entries(
    report_src_dir: str, extra_static_resources_dir: str = ""
) -> dict[str, Path]:
    """把 ``*.Report`` 目录映射成 ``Report/...`` 条目名 -> 源文件路径。"""
    src = Path(report_src_dir)
    if not src.is_dir():
        raise NotADirectoryError(f"报表源目录不存在: {src}")

    mapping: dict[str, Path] = {}
    for base in (src / "definition", src / "StaticResources"):
        if not base.is_dir():
            continue
        for file_path in sorted(base.rglob("*")):
            if file_path.is_file():
                rel = file_path.relative_to(src).as_posix()
                mapping[REPORT_PREFIX + rel] = file_path

    if extra_static_resources_dir:
        extra = Path(extra_static_resources_dir)
        if not extra.is_dir():
            raise NotADirectoryError(f"附加静态资源目录不存在: {extra}")
        for file_path in sorted(extra.rglob("*")):
            if not file_path.is_file():
                continue
            rel = PurePosixPath(file_path.relative_to(extra).as_posix())
            parts = rel.parts
            if parts and parts[0].lower() == "staticresources":
                key = REPORT_PREFIX + "/".join(parts)
            elif parts and parts[0].lower() == "registeredresources":
                key = REPORT_PREFIX + "StaticResources/" + "/".join(parts)
            else:
                key = REPORT_PREFIX + "StaticResources/RegisteredResources/" + "/".join(parts)
            mapping[key] = file_path

    return mapping


def _new_zipinfo(arcname: str, src_file: Path) -> zipfile.ZipInfo:
    mtime = time.localtime(src_file.stat().st_mtime)[:6]
    zi = zipfile.ZipInfo(arcname, date_time=mtime)
    zi.compress_type = zipfile.ZIP_DEFLATED
    zi.external_attr = 0o600 << 16
    zi.create_system = 0
    return zi


def build_pbix(
    report_src_dir: str,
    out_path: str,
    template_pbix: str,
    extra_static_resources_dir: str = "",
) -> dict:
    """用 ``template_pbix`` 当外壳，把 ``report_src_dir`` （某个 ``*.Report`` 目录）装进去。

    映射规则见模块文档。返回::

        {"ok", "out_path", "bytes", "entries_replaced", "entries_kept"}
        (+ entries_written / datamodel_bytes / datamodel_sha256 / content_types_patched)
    """
    tpl = Path(template_pbix)
    out = Path(out_path)
    if not tpl.is_file():
        return {"ok": False, "error": f"模板 pbix 不存在: {tpl}", "out_path": str(out)}
    if not Path(report_src_dir).is_dir():
        return {
            "ok": False,
            "error": f"报表源目录不存在: {report_src_dir}",
            "out_path": str(out),
        }

    new_entries = _collect_report_entries(report_src_dir, extra_static_resources_dir)
    if not new_entries:
        return {
            "ok": False,
            "error": f"{report_src_dir} 下没有 definition/ 或 StaticResources/ 内容",
            "out_path": str(out),
        }

    out.parent.mkdir(parents=True, exist_ok=True)
    tmp_out = out.with_name(out.name + ".tmp")
    if tmp_out.exists():
        tmp_out.unlink()

    replaced = 0
    kept = 0
    written = 0
    ct_patched_bytes: bytes | None = None
    ct_patched = False
    datamodel_sha = ""
    datamodel_bytes = 0

    try:
        with zipfile.ZipFile(tpl) as zin, zipfile.ZipFile(
            tmp_out, "w", zipfile.ZIP_DEFLATED, allowZip64=True
        ) as zout:
            infos = zin.infolist()

            # ---- 先算 [Content_Types].xml 是否需要补 Default ----
            raw_ct = _read_zip_entry_bytes(zin, CONTENT_TYPES_ENTRY)
            if raw_ct is not None:
                all_names = [
                    _norm(i.filename) for i in infos if not _is_report_entry(i.filename)
                ] + list(new_entries)
                patch = _ensure_content_types(raw_ct, all_names)
                if patch is not None:
                    ct_patched_bytes = patch
                    ct_patched = True

            report_written = False
            for info in infos:
                name = _norm(info.filename)
                if _is_report_entry(name):
                    replaced += 1
                    if not report_written:
                        # 就地把新的 Report/** 写入旧 Report 条目所在位置，保持条目顺序
                        for arcname in sorted(new_entries):
                            data = Path(new_entries[arcname]).read_bytes()
                            zout.writestr(_new_zipinfo(arcname, Path(new_entries[arcname])), data)
                            written += 1
                        report_written = True
                    continue  # 旧 Report 条目丢弃

                data = zin.read(info.filename)
                if name == CONTENT_TYPES_ENTRY and ct_patched_bytes is not None:
                    data = ct_patched_bytes
                if name == DATAMODEL_ENTRY:
                    datamodel_bytes = len(data)
                    datamodel_sha = _sha256(data)
                # 关键：直接复用原始 ZipInfo —— compress_type / date_time / 属性全保持
                zout.writestr(info, data)
                kept += 1

            if not report_written:  # 模板里根本没有 Report/ 条目
                for arcname in sorted(new_entries):
                    data = Path(new_entries[arcname]).read_bytes()
                    zout.writestr(_new_zipinfo(arcname, Path(new_entries[arcname])), data)
                    written += 1

        os.replace(tmp_out, out)
    except Exception as exc:  # noqa: BLE001
        if tmp_out.exists():
            try:
                tmp_out.unlink()
            except OSError:
                pass
        return {
            "ok": False,
            "error": f"{type(exc).__name__}: {exc}",
            "out_path": str(out),
        }

    return {
        "ok": True,
        "out_path": str(out),
        "bytes": out.stat().st_size,
        "entries_replaced": replaced,
        "entries_kept": kept,
        "entries_written": written,
        "datamodel_bytes": datamodel_bytes,
        "datamodel_sha256": datamodel_sha,
        "content_types_patched": ct_patched,
        "template": str(tpl),
        "report_src_dir": str(Path(report_src_dir)),
    }


# --------------------------------------------------------------------------- #
# 5) 解包 / 重新打包
# --------------------------------------------------------------------------- #
def extract_pbix(pbix_path: str, out_dir: str) -> dict:
    """把 pbix 解到目录，返回 ``{"ok", "out_dir", "files"}``（附 ``entries``）。"""
    src = Path(pbix_path)
    dst = Path(out_dir)
    if not src.is_file():
        return {"ok": False, "error": f"pbix 不存在: {src}", "out_dir": str(dst)}
    dst.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(src) as zf:
            entry_count = len(zf.infolist())
            zf.extractall(dst)
    except Exception as exc:  # noqa: BLE001
        return {
            "ok": False,
            "error": f"{type(exc).__name__}: {exc}",
            "out_dir": str(dst),
        }
    files = sum(1 for p in dst.rglob("*") if p.is_file())
    return {
        "ok": True,
        "out_dir": str(dst),
        "files": files,
        "entries": entry_count,
    }


def _read_datamodel_source(datamodel_from: str) -> bytes:
    """从 .pbix / 目录 / 裸文件里取出 DataModel 字节。"""
    src = Path(datamodel_from)
    if src.is_dir():
        direct = src / DATAMODEL_ENTRY
        if direct.is_file():
            return direct.read_bytes()
        found = sorted(p for p in src.rglob(DATAMODEL_ENTRY) if p.is_file())
        if found:
            return found[0].read_bytes()
        raise FileNotFoundError(f"目录里找不到 {DATAMODEL_ENTRY}: {src}")
    if not src.is_file():
        raise FileNotFoundError(f"datamodel_from 不存在: {src}")
    if src.suffix.lower() in (".pbix", ".pbit", ".zip"):
        with zipfile.ZipFile(src) as zf:
            return zf.read(DATAMODEL_ENTRY)
    return src.read_bytes()


def pack_dir(src_dir: str, out_path: str, datamodel_from: str = "") -> dict:
    """把一个解包目录重新打成 .pbix。

    ``datamodel_from`` 非空时（.pbix / 解包目录 / DataModel 裸文件均可），
    用它的 ``DataModel`` 条目替换目录里的那份。
    """
    src = Path(src_dir)
    out = Path(out_path)
    if not src.is_dir():
        return {"ok": False, "error": f"目录不存在: {src}", "out_path": str(out)}

    files: list[tuple[str, Path]] = []
    for p in sorted(src.rglob("*")):
        if p.is_file():
            files.append((p.relative_to(src).as_posix(), p))
    if not files:
        return {"ok": False, "error": f"目录为空: {src}", "out_path": str(out)}

    dm_bytes: bytes | None = None
    if datamodel_from:
        try:
            dm_bytes = _read_datamodel_source(datamodel_from)
        except Exception as exc:  # noqa: BLE001
            return {
                "ok": False,
                "error": f"读取 DataModel 失败: {type(exc).__name__}: {exc}",
                "out_path": str(out),
            }

    # [Content_Types].xml 按最终条目集合修补一次
    ct_names = [name for name, _ in files]
    if dm_bytes is not None and DATAMODEL_ENTRY not in ct_names:
        ct_names.append(DATAMODEL_ENTRY)

    out.parent.mkdir(parents=True, exist_ok=True)
    tmp_out = out.with_name(out.name + ".tmp")
    if tmp_out.exists():
        tmp_out.unlink()

    dm_sha = ""
    dm_size = 0
    try:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
            for name, path in files:
                data = path.read_bytes()
                if name == CONTENT_TYPES_ENTRY:
                    patch = _ensure_content_types(data, ct_names)
                    if patch is not None:
                        data = patch
                if name == DATAMODEL_ENTRY and dm_bytes is not None:
                    data = dm_bytes
                zi = _new_zipinfo(name, path)
                suffix = PurePosixPath(name).suffix.lower().lstrip(".")
                if name == DATAMODEL_ENTRY or suffix in _STORED_EXT_HINTS:
                    zi.compress_type = zipfile.ZIP_STORED  # 与 Power BI 一致：不压缩
                zout.writestr(zi, data)
                if name == DATAMODEL_ENTRY:
                    dm_sha = _sha256(data)
                    dm_size = len(data)

            if dm_bytes is not None and DATAMODEL_ENTRY not in {n for n, _ in files}:
                zi = zipfile.ZipInfo(DATAMODEL_ENTRY, date_time=time.localtime()[:6])
                zi.compress_type = zipfile.ZIP_STORED
                zout.writestr(zi, dm_bytes)
                dm_sha = _sha256(dm_bytes)
                dm_size = len(dm_bytes)

        os.replace(tmp_out, out)
    except Exception as exc:  # noqa: BLE001
        if tmp_out.exists():
            try:
                tmp_out.unlink()
            except OSError:
                pass
        return {
            "ok": False,
            "error": f"{type(exc).__name__}: {exc}",
            "out_path": str(out),
        }

    return {
        "ok": True,
        "out_path": str(out),
        "bytes": out.stat().st_size,
        "files": len(files),
        "datamodel_bytes": dm_size,
        "datamodel_sha256": dm_sha,
        "src_dir": str(src),
    }


if __name__ == "__main__":  # 手工排查用：python -m hwkit_mcp.pbipack inspect X.pbix
    import json as _json
    import sys as _sys

    try:  # 输出一律 UTF-8，避免中文在 GBK 控制台/管道里变乱码
        _sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

    if len(_sys.argv) < 3:
        print("usage: python -m hwkit_mcp.pbipack inspect|entries <pbix>", file=_sys.stderr)
        raise SystemExit(2)
    cmd, target = _sys.argv[1], _sys.argv[2]
    payload = list_entries(target) if cmd == "entries" else inspect_pbix(target)
    print(_json.dumps(payload, ensure_ascii=False, indent=2))